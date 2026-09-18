import axios from "axios";
import * as dotenv from "dotenv";
import Bottleneck from "bottleneck";

dotenv.config();

const SLAB_API_TOKEN = process.env.SLAB_API_TOKEN;
const SLAB_GRAPHQL_URL = "https://api.slab.com/v1/graphql";
const SLAB_DOMAIN = process.env.SLAB_DOMAIN || "app.slab.com";
const SLAB_RATE_LIMIT_PER_MINUTE = parseInt(
  process.env.SLAB_RATE_LIMIT_PER_MINUTE || "120"
);

const DUST_API_KEY = process.env.DUST_API_KEY;
const DUST_WORKSPACE_ID = process.env.DUST_WORKSPACE_ID;
const DUST_SPACE_ID = process.env.DUST_SPACE_ID;
const DUST_DATASOURCE_ID = process.env.DUST_DATASOURCE_ID;
const DUST_API_BASE_URL = "https://dust.tt/api/v1";
const DUST_RATE_LIMIT_PER_MINUTE = parseInt(
  process.env.DUST_RATE_LIMIT_PER_MINUTE || "120"
);

const MAX_DUST_TEXT_SIZE = 1024 * 1024; // 1MB
const SPLIT_OVERLAP = 200;

// Validate required environment variables
const requiredEnvVars = [
  "SLAB_API_TOKEN",
  "DUST_API_KEY",
  "DUST_WORKSPACE_ID",
  "DUST_SPACE_ID",
  "DUST_DATASOURCE_ID",
];

const missingVars = requiredEnvVars.filter((name) => !process.env[name]);
if (missingVars.length > 0) {
  throw new Error(
    `Missing required environment variables: ${missingVars.join(", ")}`
  );
}

const slabApi = axios.create({
  baseURL: SLAB_GRAPHQL_URL,
  headers: {
    Authorization: `Bearer ${SLAB_API_TOKEN}`,
    "Content-Type": "application/json",
  },
});

const RETRY_DELAY_MS = 2000;
slabApi.interceptors.response.use(
  (response) => response,
  async (error) => {
    if (error.response?.status === 429) {
      console.log(
        `Slab rate limit hit. Waiting ${RETRY_DELAY_MS}s before retrying...`
      );
      await new Promise((resolve) => setTimeout(resolve, RETRY_DELAY_MS));
      return slabApi.request(error.config);
    }
    if (error.response?.status === 401) {
      throw new Error(
        "Slab authentication failed. Please check your API token."
      );
    }
    return Promise.reject(error);
  }
);

const dustApi = axios.create({
  baseURL: DUST_API_BASE_URL,
  headers: {
    Authorization: `Bearer ${DUST_API_KEY}`,
    "Content-Type": "application/json",
  },
  maxContentLength: Infinity,
  maxBodyLength: Infinity,
});

const slabLimiter = new Bottleneck({
  minTime: Math.ceil(60000 / SLAB_RATE_LIMIT_PER_MINUTE),
});

const dustLimiter = new Bottleneck({
  minTime: Math.ceil(60000 / DUST_RATE_LIMIT_PER_MINUTE),
  maxConcurrent: 1,
});

const limitedSlabRequest = slabLimiter.wrap((config: any) =>
  slabApi.request(config)
);
const limitedDustRequest = dustLimiter.wrap((config: any) =>
  dustApi.request(config)
);

interface SlabTopic {
  id: string;
  name: string;
  parent?: {
    id: string;
  };
  [key: string]: any;
}

interface SlabPost {
  id: string;
  title: string;
  content?: any;
  linkAccess?: string;
  insertedAt: string;
  updatedAt: string;
  publishedAt?: string;
  archivedAt?: string;
  version?: number;
  owner?: {
    id: string;
    name: string;
    email: string;
    title?: string;
  };
  topics?: Array<{
    id: string;
    name: string;
    description?: any;
    parent?: {
      id: string;
    };
  }>;
  [key: string]: any;
}

interface DeltaOp {
  insert?: string | any;
  attributes?: any;
  delete?: number;
  retain?: number;
}

interface DeltaObject {
  ops: DeltaOp[];
}

function isDeltaOpArray(delta: unknown): delta is DeltaOp[] {
  return (
    Array.isArray(delta) &&
    delta.every((op) => typeof op === "object" && op !== null)
  );
}

function isDeltaObject(delta: unknown): delta is DeltaObject {
  return (
    typeof delta === "object" &&
    delta !== null &&
    "ops" in delta &&
    Array.isArray((delta as DeltaObject).ops)
  );
}

function toDeltaOps(delta: unknown): DeltaOp[] {
  if (!delta) {
    return [];
  }

  if (isDeltaOpArray(delta)) {
    return delta;
  }

  if (typeof delta === "string") {
    try {
      const parsed = JSON.parse(delta);
      return toDeltaOps(parsed);
    } catch {
      return [{ insert: delta }];
    }
  }

  if (isDeltaObject(delta)) {
    return delta.ops;
  }

  return [];
}

interface Heading {
  level: number;
  text: string;
  position: number;
}

interface ChunkWithContext {
  content: string;
  startHeading: Heading | null;
  endHeading: Heading | null;
  headingPath: string[];
}

interface Section {
  prefix?: string | null;
  content?: string | null;
  sections: Section[];
}

type DeltaBlockType = "heading" | "paragraph" | "list_item" | "image" | "embed";

interface DeltaBlock {
  type: DeltaBlockType;
  text: string;
  headingLevel?: number;
  headingId?: string;
  listType?: "bullet" | "ordered";
}

interface BlockSegment {
  block: DeltaBlock;
  text: string;
  byteLength: number;
}

function buildBlocksFromDelta(delta: unknown): DeltaBlock[] {
  const ops = toDeltaOps(delta);
  const blocks: DeltaBlock[] = [];
  let currentText = "";

  const flushBlock = (attributes?: any) => {
    const text = currentText.trim();
    currentText = "";
    if (!text && (!attributes || (!attributes.header && !attributes.list))) {
      return;
    }

    if (attributes?.header) {
      blocks.push({
        type: "heading",
        headingLevel: Number(attributes.header) || 1,
        headingId: attributes["header-id"],
        text,
      });
    } else if (attributes?.list) {
      blocks.push({
        type: "list_item",
        listType: attributes.list === "ordered" ? "ordered" : "bullet",
        text,
      });
    } else {
      blocks.push({
        type: "paragraph",
        text,
      });
    }
  };

  for (const op of ops) {
    if (op.insert == null) {
      continue;
    }

    if (typeof op.insert === "string") {
      let remaining = op.insert;
      while (remaining.length > 0) {
        const newlineIdx = remaining.indexOf("\n");
        if (newlineIdx === -1) {
          currentText += remaining;
          remaining = "";
        } else {
          currentText += remaining.slice(0, newlineIdx);
          flushBlock(op.attributes);
          remaining = remaining.slice(newlineIdx + 1);
        }
      }
    } else if (typeof op.insert === "object") {
      if (currentText.trim().length > 0) {
        flushBlock();
      } else {
        currentText = "";
      }

      if (op.insert.image) {
        const images = Array.isArray(op.insert.image)
          ? op.insert.image
          : [op.insert.image];
        for (const imageEntry of images) {
          const source =
            typeof imageEntry === "string"
              ? imageEntry
              : imageEntry?.source || "image";
          blocks.push({
            type: "image",
            text: `[Image: ${source}]`,
          });
        }
      } else {
        blocks.push({
          type: "embed",
          text: "[Embedded content]",
        });
      }
    }
  }

  if (currentText.trim().length > 0) {
    flushBlock();
  }

  return blocks;
}

function blocksToTextSegments(blocks: DeltaBlock[]): BlockSegment[] {
  const segments: BlockSegment[] = [];
  let orderedCounter = 1;
  let previousListType: "bullet" | "ordered" | null = null;

  for (const block of blocks) {
    let text = "";
    switch (block.type) {
      case "heading": {
        text = `${block.text}\n`;
        orderedCounter = 1;
        previousListType = null;
        break;
      }
      case "list_item": {
        if (block.listType === "ordered") {
          if (previousListType !== "ordered") {
            orderedCounter = 1;
          }
          text = `${orderedCounter}. ${block.text}\n`;
          orderedCounter += 1;
          previousListType = "ordered";
        } else {
          text = `- ${block.text}\n`;
          previousListType = "bullet";
        }
        break;
      }
      case "image":
      case "embed": {
        text = `${block.text}\n`;
        orderedCounter = 1;
        previousListType = null;
        break;
      }
      default: {
        text = `${block.text}\n\n`;
        orderedCounter = 1;
        previousListType = null;
      }
    }

    segments.push({
      block,
      text,
      byteLength: Buffer.byteLength(text, "utf8"),
    });
  }

  return segments;
}

function findHeadingInRange(
  segments: BlockSegment[],
  blockPositions: number[],
  startIdx: number,
  endIdx: number,
  fromStart: boolean
): Heading | null {
  if (fromStart) {
    for (let i = startIdx; i < endIdx; i++) {
      if (segments[i].block.type === "heading") {
        return {
          level: segments[i].block.headingLevel || 1,
          text: segments[i].block.text,
          position: blockPositions[i],
        };
      }
    }
  } else {
    for (let i = endIdx - 1; i >= startIdx; i--) {
      if (segments[i].block.type === "heading") {
        return {
          level: segments[i].block.headingLevel || 1,
          text: segments[i].block.text,
          position: blockPositions[i],
        };
      }
    }
  }
  return null;
}

function splitContentForDust(
  delta: unknown,
  maxSize: number,
  overlap: number
): ChunkWithContext[] {
  const blocks = buildBlocksFromDelta(delta);
  if (blocks.length === 0) {
    return [];
  }

  const segments = blocksToTextSegments(blocks);
  if (segments.length === 0) {
    return [];
  }

  return chunkBlockSegments(segments, maxSize, overlap);
}

function chunkBlockSegments(
  segments: BlockSegment[],
  maxSize: number,
  overlap: number
): ChunkWithContext[] {
  if (segments.length === 0) {
    return [];
  }

  const total = segments.length;
  const blockPositions: number[] = new Array(total);
  const previousHeadingBefore: (Heading | null)[] = new Array(total + 1);
  const headingPathBefore: string[][] = new Array(total);

  // Precompute cumulative byte positions for each segment

  let cumulativeBytes = 0;
  for (let i = 0; i < total; i++) {
    blockPositions[i] = cumulativeBytes;
    cumulativeBytes += segments[i].byteLength;
  }

  const headingStack: { level: number; text: string; position: number }[] = [];

  // Build heading hierarchy: track path and previous heading at each position
  for (let i = 0; i < total; i++) {
    headingPathBefore[i] = headingStack.map((h) => h.text);
    previousHeadingBefore[i] =
      headingStack.length > 0
        ? {
            level: headingStack[headingStack.length - 1].level,
            text: headingStack[headingStack.length - 1].text,
            position: headingStack[headingStack.length - 1].position,
          }
        : null;

    if (segments[i].block.type === "heading") {
      const level = segments[i].block.headingLevel || 1;
      // Can remove everything at this heading level and below
      headingStack.splice(level - 1);
      headingStack[level - 1] = {
        level,
        text: segments[i].block.text,
        position: blockPositions[i],
      };
    }
  }

  previousHeadingBefore[total] =
    headingStack.length > 0
      ? {
          level: headingStack[headingStack.length - 1].level,
          text: headingStack[headingStack.length - 1].text,
          position: headingStack[headingStack.length - 1].position,
        }
      : null;

  const chunks: ChunkWithContext[] = [];
  let startIdx = 0;
  const overlapBytes = Math.min(overlap, Math.floor(maxSize / 2));

  while (startIdx < total) {
    let endIdx = startIdx;
    let chunkBytes = 0;
    const chunkParts: string[] = [];

    // Accumulate segments until we reach maxSize
    while (endIdx < total) {
      const nextLength = segments[endIdx].byteLength;

      // Stop if adding this segment would exceed maxSize (and we already have content)
      if (chunkBytes + nextLength > maxSize && chunkBytes > 0) {
        break;
      }

      // Handle oversized segment: a single segment that exceeds maxSize
      // Split it at byte boundaries to fit within the current chunk capacity
      if (nextLength > maxSize) {
        const remainingInChunk = maxSize - chunkBytes;
        const splitSize = remainingInChunk > 0 ? remainingInChunk : maxSize;
        const text = segments[endIdx].text;
        const buffer = Buffer.from(text, "utf8");
        const firstPart = buffer.subarray(0, splitSize).toString("utf8");
        const remainingPart = buffer.subarray(splitSize).toString("utf8");

        // Add the first part to current chunk
        if (firstPart.length > 0) {
          chunkParts.push(firstPart);
          chunkBytes += Buffer.byteLength(firstPart, "utf8");
        }

        // Update the segment with remaining part so it gets processed in next iteration
        if (remainingPart.length > 0) {
          segments[endIdx].text = remainingPart;
          segments[endIdx].byteLength = Buffer.byteLength(
            remainingPart,
            "utf8"
          );
        } else {
          endIdx++;
        }
        continue;
      }

      // Normal case: add the entire segment to the chunk
      chunkParts.push(segments[endIdx].text);
      chunkBytes += nextLength;
      endIdx++;

      // Stop if we've reached maxSize
      if (chunkBytes >= maxSize) {
        break;
      }
    }

    if (chunkParts.length === 0) {
      startIdx = endIdx;
      continue;
    }

    // Build the chunk with heading context
    const chunkContent = chunkParts.join("");
    // Find the first heading in this chunk, or use the previous heading as context
    const startHeading =
      findHeadingInRange(segments, blockPositions, startIdx, endIdx, true) ||
      previousHeadingBefore[startIdx];
    // Find the last heading in this chunk, or use the previous heading as context
    const endHeading =
      findHeadingInRange(segments, blockPositions, startIdx, endIdx, false) ||
      previousHeadingBefore[endIdx] ||
      previousHeadingBefore[startIdx];
    // Get the full heading path (hierarchy) at the start of this chunk
    const headingPath =
      headingPathBefore[startIdx] && headingPathBefore[startIdx].length > 0
        ? [...headingPathBefore[startIdx]]
        : [];

    chunks.push({
      content: chunkContent,
      startHeading,
      endHeading,
      headingPath,
    });

    // If we've processed all segments, we're done
    if (endIdx >= total) {
      break;
    }

    // Calculate overlap for next chunk
    // Overlap ensures context is preserved between chunks by including some content
    // from the end of the current chunk at the start of the next chunk
    if (overlapBytes <= 0) {
      startIdx = endIdx;
      continue;
    }

    // Work backwards from endIdx to find where to start the next chunk
    // We want to keep approximately overlapBytes worth of content
    let bytesToKeep = overlapBytes;
    let newStartIdx = endIdx;
    while (newStartIdx > startIdx && bytesToKeep > 0) {
      newStartIdx--;
      bytesToKeep -= segments[newStartIdx].byteLength;
    }

    // If we couldn't find enough content to overlap, just move to endIdx
    // Otherwise, start the next chunk at newStartIdx to include the overlap
    if (newStartIdx <= startIdx) {
      startIdx = endIdx;
    } else {
      startIdx = newStartIdx;
    }
  }

  return chunks;
}

async function* getAllSlabPosts(): AsyncGenerator<SlabPost, void, unknown> {
  console.log("Fetching Slab posts...");

  try {
    let allPostIds: string[] = [];

    const orgQuery = `
      query GetOrganizationPosts {
        organization {
          posts {
            id
          }
        }
      }
    `;

    const orgResponse = await limitedSlabRequest({
      method: "POST",
      data: {
        query: orgQuery,
      },
    });

    if (
      !orgResponse.data?.errors &&
      orgResponse.data?.data?.organization?.posts
    ) {
      allPostIds = orgResponse.data.data.organization.posts.map(
        (p: any) => p.id
      );
      console.log(`Found ${allPostIds.length} posts via organization.posts`);
    }

    if (allPostIds.length === 0) {
      console.log("No posts found");
      return;
    }

    // Now batch fetch posts in chunks of 100 (maxItems constraint)
    const BATCH_SIZE = 100;
    let batchStart = 0;

    while (batchStart < allPostIds.length) {
      const batchIds = allPostIds.slice(batchStart, batchStart + BATCH_SIZE);
      const batchEnd = Math.min(batchStart + BATCH_SIZE, allPostIds.length);

      const postsQuery = `
        query GetPosts($ids: [ID!]!) {
          posts(ids: $ids) {
            id
            title
            content
            linkAccess
            insertedAt
            updatedAt
            publishedAt
            archivedAt
            version
            owner {
              id
              name
              email
              title
            }
            topics {
              id
              name
              description
              parent {
                id
              }
            }
          }
        }
      `;

      const response = await limitedSlabRequest({
        method: "POST",
        data: {
          query: postsQuery,
          variables: { ids: batchIds },
        },
      });

      if (response.data?.errors) {
        throw new Error(
          `GraphQL query failed: ${response.data.errors
            .map((e: any) => e.message)
            .join(", ")}`
        );
      }

      const posts = response.data?.data?.posts || [];

      for (const post of posts) {
        yield post;
      }

      console.log(
        `Retrieved batch ${Math.floor(batchStart / BATCH_SIZE) + 1}: ${
          posts.length
        } posts`
      );

      batchStart = batchEnd;
    }
  } catch (error: any) {
    if (error.response?.data) {
      console.error(
        "API Response:",
        JSON.stringify(error.response.data, null, 2)
      );
    }
    console.error(`Error fetching posts: ${error.message}`);
    if (error.response?.data?.errors) {
      console.error(
        "GraphQL errors:",
        JSON.stringify(error.response.data.errors, null, 2)
      );
    }
    throw error;
  }
}

async function getTopicDetails(topicId: string): Promise<SlabTopic | null> {
  try {
    const query = `
      query GetTopic($id: ID!) {
        topic(id: $id) {
          id
          name
          parent {
            id
          }
        }
      }
    `;

    const response = await limitedSlabRequest({
      method: "POST",
      data: {
        query,
        variables: { id: topicId },
      },
    });

    if (response.data?.data?.topic) {
      return response.data.data.topic;
    } else if (response.data?.topic) {
      return response.data.topic;
    }
    return null;
  } catch (error: any) {
    console.error(`Error fetching topic ${topicId}: ${error.message}`);
    return null;
  }
}

function handleDustApiError(
  error: any,
  postTitle: string,
  partNumber?: number
): never {
  const partSuffix = partNumber && partNumber > 1 ? ` part ${partNumber}` : "";
  if (error.response) {
    console.error(`Dust API error for post "${postTitle}"${partSuffix}:`, {
      status: error.response.status,
      statusText: error.response.statusText,
      url: error.config?.url,
      data: error.response.data,
    });
  } else {
    console.error(
      `Dust API error for post "${postTitle}"${partSuffix}:`,
      error.message
    );
  }
  throw error;
}

async function buildTopicHierarchyPath(
  topicId: string | undefined
): Promise<string> {
  if (!topicId) {
    return "";
  }

  const topicPath: string[] = [];
  let currentTopicId: string | undefined = topicId;

  while (currentTopicId) {
    const topic = await getTopicDetails(currentTopicId);
    if (!topic) {
      break;
    }

    topicPath.unshift(topic.name); // Add to beginning to build path from root

    // Check if topic has a parent
    if (topic.parent?.id) {
      currentTopicId = topic.parent.id;
    } else {
      break;
    }
  }

  return topicPath.join(" > ");
}

async function upsertToDustDatasource(post: SlabPost) {
  const baseDocumentId = `slab-${post.id}`;

  // Build topic hierarchy path from first topic if available
  const firstTopic =
    post.topics && post.topics.length > 0 ? post.topics[0] : null;
  const topicHierarchyPath = await buildTopicHierarchyPath(firstTopic?.id);

  const baseMetadataLines = [
    `Title: ${post.title}`,
    `Post ID: ${post.id}`,
    `Created At: ${post.insertedAt}`,
    `Updated At: ${post.updatedAt}`,
    ...(firstTopic
      ? [
          topicHierarchyPath
            ? `Topic Path: ${topicHierarchyPath}`
            : `Topic: ${firstTopic.name}`,
        ]
      : []),
  ];

  // Calculate max content size accounting for part metadata overhead
  const maxHierarchyOverhead = 300;
  const maxSizeOfPartMeta = `\n(Part 999 of 999)`;
  const maxPartMetaOverhead = Buffer.byteLength(maxSizeOfPartMeta, "utf8");

  // Base metadata overhead (will be recalculated per chunk with hierarchy)
  const baseMetadataText = `---\n${baseMetadataLines.join("\n")}\n---\n\n`;
  const baseMetadataOverhead = Buffer.byteLength(baseMetadataText, "utf8");
  const maxContentSize =
    MAX_DUST_TEXT_SIZE -
    baseMetadataOverhead -
    maxPartMetaOverhead -
    maxHierarchyOverhead;

  if (maxContentSize <= 0) {
    throw new Error(
      `Metadata overhead is too large for post "${post.title}". Cannot split content.`
    );
  }

  const contentChunks = splitContentForDust(
    post.content,
    maxContentSize,
    SPLIT_OVERLAP
  );

  if (contentChunks.length === 0) {
    // No chunks - create empty document with just metadata
    const metadataText = `---\n${baseMetadataLines.join("\n")}\n---\n\n`;
    const section: Section = {
      prefix: post.title,
      content: metadataText,
      sections: [],
    };

    try {
      await limitedDustRequest({
        method: "POST",
        url: `/w/${DUST_WORKSPACE_ID}/spaces/${DUST_SPACE_ID}/data_sources/${DUST_DATASOURCE_ID}/documents/${baseDocumentId}`,
        data: {
          section,
          title: post.title,
          source_url: `https://${SLAB_DOMAIN}/posts/${post.id}`,
        },
      });
    } catch (error: any) {
      handleDustApiError(error, post.title);
    }
    return;
  }

  // Process chunks incrementally, building sections and uploading documents as we go
  let currentBatch: Section[] = [];
  let documentPartNumber = 1;

  async function uploadBatch(batch: Section[], partNumber: number) {
    const batchSection: Section = {
      prefix: post.title,
      content: batch.length === 1 ? batch[0].content : null,
      sections: batch.length === 1 ? [] : batch,
    };
    const documentId =
      partNumber === 1 ? baseDocumentId : `${baseDocumentId}-part${partNumber}`;
    const title =
      contentChunks.length === 1
        ? post.title
        : `${post.title} (Part ${partNumber})`;

    try {
      await limitedDustRequest({
        method: "POST",
        url: `/w/${DUST_WORKSPACE_ID}/spaces/${DUST_SPACE_ID}/data_sources/${DUST_DATASOURCE_ID}/documents/${documentId}`,
        data: {
          section: batchSection,
          title,
          source_url: `https://${SLAB_DOMAIN}/posts/${post.id}`,
        },
      });
    } catch (error: any) {
      handleDustApiError(error, post.title, partNumber);
    }
  }

  for (let i = 0; i < contentChunks.length; i++) {
    const chunk = contentChunks[i];
    const chunkMetadataLines = [...baseMetadataLines];

    // Add section path in metadata
    if (chunk.headingPath.length > 0) {
      chunkMetadataLines.push(`Section Path: ${chunk.headingPath.join(" > ")}`);
    }

    // Add part information
    chunkMetadataLines.push(`Part: ${i + 1} of ${contentChunks.length}`);

    const chunkMetadataText = `---\n${chunkMetadataLines.join("\n")}\n---\n\n`;
    const partMeta = `\n(Part ${i + 1} of ${contentChunks.length})`;
    const content = chunkMetadataText + chunk.content + partMeta;

    // Verify the final size doesn't exceed the limit
    const finalSize = Buffer.byteLength(content, "utf8");
    if (finalSize > MAX_DUST_TEXT_SIZE) {
      console.warn(
        `[SKIP] Part ${i + 1} of "${post.title}" exceeds Dust limit. Skipping.`
      );
      continue;
    }

    // Build section prefix from heading path
    let sectionPrefix: string;
    if (chunk.headingPath.length > 0) {
      sectionPrefix = chunk.headingPath[chunk.headingPath.length - 1];
    } else if (chunk.startHeading) {
      sectionPrefix = chunk.startHeading.text;
    } else {
      sectionPrefix = `Part ${i + 1}`;
    }

    const newSection: Section = {
      prefix: sectionPrefix,
      content: content,
      sections: [],
    };

    // Test if adding this section would exceed the limit
    const testBatch: Section[] = [...currentBatch, newSection];
    const testBatchSection: Section = {
      prefix: post.title,
      content: testBatch.length === 1 ? testBatch[0].content : null,
      sections: testBatch.length === 1 ? [] : testBatch,
    };
    const testTitle =
      contentChunks.length === 1
        ? post.title
        : `${post.title} (Part ${documentPartNumber})`;
    const testPayload = {
      section: testBatchSection,
      title: testTitle,
      source_url: `https://${SLAB_DOMAIN}/posts/${post.id}`,
    };
    const testPayloadSize = Buffer.byteLength(
      JSON.stringify(testPayload),
      "utf8"
    );

    if (testPayloadSize > MAX_DUST_TEXT_SIZE && currentBatch.length > 0) {
      await uploadBatch(currentBatch, documentPartNumber);
      // Start new batch with current section
      documentPartNumber++;
    } else {
      // Add to current batch
      currentBatch.push(newSection);
    }
  }

  if (currentBatch.length > 0) {
    await uploadBatch(currentBatch, documentPartNumber);
  }
}

async function main() {
  try {
    let processed = 0;
    let errors = 0;

    for await (const post of getAllSlabPosts()) {
      try {
        await upsertToDustDatasource(post);

        processed++;
        console.log(
          `✓ Successfully processed: ${post.title} (${processed} processed, ${errors} errors)`
        );
      } catch (error: any) {
        errors++;
        console.error(`✗ Error processing "${post.title}": ${error.message}`);
        console.error(`   Skipping and continuing with next document...`);
      }
    }

    console.log(`Processed: ${processed} posts`);
    console.log(`Errors: ${errors} posts`);
  } catch (error: any) {
    console.error("Synchronization failed:", error.message);
    process.exit(1);
  }
}

main();
