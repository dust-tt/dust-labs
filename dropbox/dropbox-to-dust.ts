import axios, { AxiosResponse } from 'axios';
import * as dotenv from 'dotenv';
import Bottleneck from 'bottleneck';
import path from 'path';

dotenv.config();

// Dropbox API
const DROPBOX_API_KEY = process.env.DROPBOX_API_KEY;
const DROPBOX_ROOT_PATH = process.env.DROPBOX_ROOT_PATH || '';
const DROPBOX_EXTENSION_FILTER = process.env.DROPBOX_EXTENSION_FILTER;
const DROPBOX_MAX_CONCURRENT = 2; // tuned for rate limits

// Dust API
const DUST_API_KEY = process.env.DUST_API_KEY;
const DUST_WORKSPACE_ID = process.env.DUST_WORKSPACE_ID;
const DUST_SPACE_ID = process.env.DUST_SPACE_ID;
const DUST_DATASOURCE_ID = process.env.DUST_DATASOURCE_ID;
const DUST_RATE_LIMIT = parseInt(process.env.DUST_RATE_LIMIT || '120');

const MAX_DUST_TEXT_SIZE = 1024 * 1024; // 1MB
const SPLIT_OVERLAP = 200; // chars of overlap between splits for context

// CLI arg for extension filter
const extArgIdx = process.argv.indexOf('--ext');
const EXTENSION = extArgIdx !== -1 && process.argv[extArgIdx + 1] ? process.argv[extArgIdx + 1] : DROPBOX_EXTENSION_FILTER;

// Validate required env vars
const missingEnvVars = [
  ['DROPBOX_API_KEY', DROPBOX_API_KEY],
  ['DUST_API_KEY', DUST_API_KEY],
  ['DUST_WORKSPACE_ID', DUST_WORKSPACE_ID],
  ['DUST_SPACE_ID', DUST_SPACE_ID],
  ['DUST_DATASOURCE_ID', DUST_DATASOURCE_ID],
].filter(([name, value]) => !value).map(([name]) => name);

if (missingEnvVars.length > 0) {
  throw new Error(`Please provide values for the following environment variables in the .env file: ${missingEnvVars.join(', ')}`);
}

// Dropbox API client (metadata)
const dropboxApi = axios.create({
  baseURL: 'https://api.dropboxapi.com/2',
  headers: {
    'Authorization': `Bearer ${DROPBOX_API_KEY}`,
    'Content-Type': 'application/json',
  },
});

// Dropbox Content API client (for export/download)
const dropboxContentApi = axios.create({
  baseURL: 'https://content.dropboxapi.com/2',
  headers: {
    'Authorization': `Bearer ${DROPBOX_API_KEY}`,
  },
  responseType: 'text',
});

// Add Dropbox 429 rate limit handling
const RETRY_DELAY_MS = 2000;
dropboxContentApi.interceptors.response.use(
  (response) => response,
  async (error) => {
    if (error.response && error.response.status === 429) {
      console.error('Dropbox rate limit hit. Backing off for 2s and retrying...');
      await new Promise((resolve) => setTimeout(resolve, RETRY_DELAY_MS));
      return dropboxContentApi.request(error.config);
    }
    return Promise.reject(error);
  }
);

dropboxApi.interceptors.response.use(
  (response) => response,
  async (error) => {
    if (error.response && error.response.status === 429) {
      console.error('Dropbox rate limit hit (metadata). Backing off for 2s and retrying...');
      await new Promise((resolve) => setTimeout(resolve, RETRY_DELAY_MS));
      return dropboxApi.request(error.config);
    }
    return Promise.reject(error);
  }
);

// Dust API client
const dustApi = axios.create({
  baseURL: 'https://dust.tt/api/v1',
  headers: {
    'Authorization': `Bearer ${DUST_API_KEY}`,
    'Content-Type': 'application/json',
  },
  maxContentLength: Infinity,
  maxBodyLength: Infinity,
});

// Rate limiters
// This allows up to (1000ms / 150ms) * 2 ≈ 13 requests/second (theoretical max, but concurrency is usually limited by I/O)
const DROPBOX_REQUESTS_PER_SECOND = 6; // Target requests per second
const dropboxLimiter = new Bottleneck({
  maxConcurrent: DROPBOX_MAX_CONCURRENT,
  minTime: Math.ceil(1000 / DROPBOX_REQUESTS_PER_SECOND),
});
const dustLimiter = new Bottleneck({
  minTime: Math.ceil(60000 / DUST_RATE_LIMIT),
  maxConcurrent: 1,
});

const limitedDropboxRequest = dropboxLimiter.wrap((config: any) => dropboxApi.request(config));
const limitedDropboxContentRequest = dropboxLimiter.wrap((config: any) => dropboxContentApi.request(config));
const limitedDustRequest = dustLimiter.wrap((config: any) => dustApi.request(config));

interface DropboxFileMetadata {
  id: string;
  name: string;
  path_display: string;
  client_modified: string;
  server_modified: string;
  is_paper_doc?: boolean;
  [key: string]: any;
}

async function listAllFiles(folderPath: string, extension: string | undefined): Promise<DropboxFileMetadata[]> {
  let hasMore = true;
  let cursor: string | null = null;
  let files: DropboxFileMetadata[] = [];

  while (hasMore) {
    let res: AxiosResponse<any>;
    if (!cursor) {
      res = await limitedDropboxRequest({
        method: 'POST',
        url: '/files/list_folder',
        data: {
          path: folderPath,
          recursive: true,
          include_media_info: false,
          include_deleted: false,
          include_has_explicit_shared_members: false,
          include_mounted_folders: true,
        },
      });
    } else {
      res = await limitedDropboxRequest({
        method: 'POST',
        url: '/files/list_folder/continue',
        data: { cursor },
      });
    }
    let entries = res.data.entries.filter((entry: any) => entry['.tag'] === 'file');
    if (extension) {
      entries = entries.filter((entry: any) => entry.name.endsWith(extension));
    }
    files = files.concat(entries);
    hasMore = res.data.has_more;
    cursor = res.data.cursor;
  }
  return files;
}

async function exportDropboxFile(file: DropboxFileMetadata): Promise<{ content: string; format: string }> {
  // Match .paper, .papert, etc.
  const isPaperDoc = file.name.endsWith('.paper') || file.name.endsWith('.papert');
  const contentHeaders = {
    'Authorization': `Bearer ${DROPBOX_API_KEY}`,
    'Dropbox-API-Arg': JSON.stringify(
      isPaperDoc
        ? { path: file.path_display, export_format: 'markdown' }
        : { path: file.path_display }
    ),
    'Content-Type': 'text/plain',
  };

  if (isPaperDoc) {
    // Use Dropbox Content API for export
    const res = await limitedDropboxContentRequest({
      method: 'POST',
      url: '/files/export',
      headers: contentHeaders,
      data: null,
      responseType: 'text',
    });
    return { content: res.data, format: 'markdown' };
  } else {
    // For other files, use /files/download via Content API
    const res = await limitedDropboxContentRequest({
      method: 'POST',
      url: '/files/download',
      headers: contentHeaders,
      data: null,
      responseType: 'text',
    });
    return { content: res.data, format: path.extname(file.name).replace('.', '') };
  }
}

function splitContentForDust(text: string, maxSize: number, overlap: number): string[] {
  const parts: string[] = [];
  let start = 0;
  while (start < text.length) {
    let end = start + maxSize;
    if (end > text.length) end = text.length;
    let part = text.slice(start, end);
    parts.push(part);
    if (end === text.length) break;
    start = end - overlap; // overlap for context
    if (start < 0) start = 0;
  }
  return parts;
}

async function upsertToDustDatasource(file: DropboxFileMetadata, content: string, format: string) {
  const baseDocumentId = `${file.name}`;
  const metadata = [
    `File Name: ${file.name}`,
    `Path: ${file.path_display}`,
    `File ID: ${file.id}`,
    `Client Modified: ${file.client_modified}`,
    `Server Modified: ${file.server_modified}`,
    `Format: ${format}`,
  ].join('\n');

  const baseText = `---\n${metadata}\n---\n\n`;
  const maxContentSize = MAX_DUST_TEXT_SIZE - Buffer.byteLength(baseText, 'utf8');
  const contentParts = splitContentForDust(content, maxContentSize, SPLIT_OVERLAP);

  for (let i = 0; i < contentParts.length; i++) {
    const partText = baseText + contentParts[i];
    const documentId = contentParts.length === 1 ? baseDocumentId : `${baseDocumentId}-part${i + 1}`;
    const partMeta = contentParts.length === 1 ? '' : `\n(Part ${i + 1} of ${contentParts.length})`;
    const textWithPart = partText + partMeta;
    if (Buffer.byteLength(textWithPart, 'utf8') > MAX_DUST_TEXT_SIZE) {
      console.warn(`[SKIP] Even split part for ${file.name} exceeds Dust 1MB text limit. Skipping part ${i + 1}.`);
      continue;
    }
    const url = `/w/${DUST_WORKSPACE_ID}/spaces/${DUST_SPACE_ID}/data_sources/${DUST_DATASOURCE_ID}/documents/${documentId}`;
    await limitedDustRequest({
      method: 'POST',
      url,
      data: {
        text: textWithPart,
        source_url: `https://www.dropbox.com/home${file.path_display}`,
      },
    });
    console.log(`[UPLOADED] ${file.name} part ${i + 1}/${contentParts.length} as ${documentId}`);
  }
}

async function main() {
  console.log(`Syncing Dropbox files${EXTENSION ? ` with extension '${EXTENSION}'` : ''} to Dust datasource...`);
  try {
    const files = await listAllFiles(DROPBOX_ROOT_PATH, EXTENSION || undefined);
    console.log(`Found ${files.length} files${EXTENSION ? ` with extension '${EXTENSION}'` : ''}.`);
    let processed = 0;
    for (const file of files) {
      try {
        const { content, format } = await exportDropboxFile(file);
        await upsertToDustDatasource(file, content, format);
        processed++;
        console.log(`[${processed}/${files.length}] Synced: ${file.name}`);
      } catch (err) {
        console.error(`Error processing file ${file.name}:`, err);
      }
    }
    console.log('All files processed.');
  } catch (err) {
    console.error('An error occurred:', err);
  }
}

main(); 