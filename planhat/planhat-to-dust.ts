import axios, { AxiosInstance } from "axios";
import Bottleneck from "bottleneck";
import dotenv from "dotenv";
import { Worker, isMainThread, parentPort, workerData } from "worker_threads";
import { fileURLToPath } from "url";

dotenv.config();

const PLANHAT_API_TOKEN = process.env.PLANHAT_API_TOKEN;
const DUST_API_KEY = process.env.DUST_API_KEY;
const DUST_WORKSPACE_ID = process.env.DUST_WORKSPACE_ID;
const DUST_SPACE_ID = process.env.DUST_SPACE_ID;
const DUST_DATASOURCE_ID = process.env.DUST_DATASOURCE_ID;
const DUST_REGION = process.env.DUST_REGION || "US";
const LOOKBACK_DAYS = parseInt(process.env.LOOKBACK_DAYS || "7", 10);
const THREADS_NUMBER = parseInt(process.env.THREADS_NUMBER || "1", 10);

if (
  !PLANHAT_API_TOKEN ||
  !DUST_API_KEY ||
  !DUST_WORKSPACE_ID ||
  !DUST_SPACE_ID ||
  !DUST_DATASOURCE_ID
) {
  console.error("Missing required environment variables");
  process.exit(1);
}

const planhatApi = axios.create({
  baseURL: "https://api.planhat.com",
  headers: {
    Authorization: `Bearer ${PLANHAT_API_TOKEN}`,
    "Content-Type": "application/json",
  },
});

// Determine Dust API base URL based on region
const dustBaseUrl =
  DUST_REGION === "EU" ? "https://eu.dust.tt/api/v1" : "https://dust.tt/api/v1";

const dustApi = axios.create({
  baseURL: dustBaseUrl,
  headers: {
    Authorization: `Bearer ${DUST_API_KEY}`,
    "Content-Type": "application/json",
  },
});

const planhatLimiter = new Bottleneck({
  maxConcurrent: 1,
  minTime: 200, // 5 requests per second max
});

const dustLimiter = new Bottleneck({
  maxConcurrent: 1,
  minTime: 500 / THREADS_NUMBER,
});

// NEW: helper to log Planhat API errors on a single line
function logPlanhatError(resource: string, context: string, error: unknown) {
  const err = error as any;
  const status = err?.response?.status ?? "";
  const statusText = err?.response?.statusText ?? "";
  const message =
    err?.response?.data?.message ?? err?.message ?? "Unknown error";
  console.error(
    `Planhat API error while fetching ${resource} for ${context}: ${status} ${statusText} - ${message}`
  );
}

interface PlanhatCompany {
  _id: string;
  name: string;
  slug?: string;
  custom?: Record<string, any>;
  mrr?: number;
  status?: string;
  health?: string;
  renewal?: string;
  phase?: string;
  owner?: {
    name?: string;
    email?: string;
  };
  createdAt?: string;
  updatedAt?: string;
  [key: string]: any;
}

interface PlanhatConversation {
  _id: string;
  companyId: string;
  subject?: string;
  type?: string;
  date?: string;
  description?: string;
  [key: string]: any;
}

interface PlanhatEnduser {
  _id: string;
  companyId: string;
  name?: string;
  email?: string;
  role?: string;
  status?: string;
  [key: string]: any;
}

interface PlanhatNPS {
  _id: string;
  companyId: string;
  score?: number;
  comment?: string;
  date?: string;
  enduser?: {
    name?: string;
    email?: string;
  };
  [key: string]: any;
}

interface PlanhatProject {
  _id: string;
  companyId: string;
  name?: string;
  status?: string;
  phase?: string;
  startDate?: string;
  endDate?: string;
  [key: string]: any;
}

interface PlanhatAsset {
  _id: string;
  companyId: string;
  name?: string;
  type?: string;
  status?: string;
  value?: number;
  [key: string]: any;
}

interface DustSection {
  prefix: string;
  content: string | null;
  sections: DustSection[];
}

interface DustDocument {
  dataSourceId: string;
  documentId: string;
  title: string;
  section: DustSection;
  parents: string[];
  tags: string[];
  sourceUrl?: string;
  timestampMs?: number;
}

function sanitizeDocumentId(id: string): string {
  return id.replace(/[^a-zA-Z0-9-_]/g, "_").substring(0, 128);
}

async function fetchCompanies(updatedSince?: Date): Promise<PlanhatCompany[]> {
  const allCompanies: PlanhatCompany[] = [];
  let skip = 0;
  const limit = 50;

  try {
    while (true) {
      const params: any = { limit, skip };
      if (updatedSince) {
        params.updatedAfter = updatedSince.toISOString();
      }

      const response = await planhatLimiter.schedule(() =>
        planhatApi.get("/companies", { params })
      );

      const companies: PlanhatCompany[] = response.data || [];
      allCompanies.push(...companies);

      if (companies.length < limit) {
        break; // No more pages
      }
      skip += limit;
    }
    return allCompanies;
  } catch (error) {
    logPlanhatError("companies", updatedSince?.toISOString() || "all", error);
    throw error;
  }
}

async function fetchConversations(
  companyId: string
): Promise<PlanhatConversation[]> {
  try {
    const response = await planhatLimiter.schedule(() =>
      planhatApi.get(`/conversations`, { params: { companyId } })
    );
    return response.data || [];
  } catch (error) {
    logPlanhatError("conversations", companyId, error);
    return [];
  }
}

async function fetchEndusers(companyId: string): Promise<PlanhatEnduser[]> {
  try {
    const response = await planhatLimiter.schedule(() =>
      planhatApi.get(`/endusers`, { params: { companyId } })
    );
    return response.data || [];
  } catch (error) {
    logPlanhatError("endusers", companyId, error);
    return [];
  }
}

async function fetchNPS(companyId: string): Promise<PlanhatNPS[]> {
  try {
    const response = await planhatLimiter.schedule(() =>
      planhatApi.get(`/nps`, { params: { companyId } })
    );
    return response.data || [];
  } catch (error) {
    logPlanhatError("nps", companyId, error);
    return [];
  }
}

async function fetchProjects(companyId: string): Promise<PlanhatProject[]> {
  try {
    const response = await planhatLimiter.schedule(() =>
      planhatApi.get(`/projects`, { params: { companyId } })
    );
    return response.data || [];
  } catch (error) {
    logPlanhatError("projects", companyId, error);
    return [];
  }
}

async function fetchAssets(companyId: string): Promise<PlanhatAsset[]> {
  try {
    const response = await planhatLimiter.schedule(() =>
      planhatApi.get(`/assets`, { params: { companyId } })
    );
    return response.data || [];
  } catch (error) {
    logPlanhatError("assets", companyId, error);
    return [];
  }
}

function buildCompanySection(company: PlanhatCompany): DustSection {
  const sections: DustSection[] = [];

  // Basic information
  sections.push({
    prefix: "basic_information",
    content: `Company Name: ${company.name}
ID: ${company._id}
Status: ${company.status || "N/A"}
Health: ${company.health || "N/A"}
MRR: ${company.mrr || "N/A"}
Phase: ${company.phase || "N/A"}
Renewal: ${company.renewal || "N/A"}
Owner: ${company.owner?.name || "N/A"} (${company.owner?.email || "N/A"})
Created: ${company.createdAt || "N/A"}
Updated: ${company.updatedAt || "N/A"}`,
    sections: [],
  });

  // Custom fields
  if (company.custom && Object.keys(company.custom).length > 0) {
    sections.push({
      prefix: "custom_fields",
      content: Object.entries(company.custom)
        .map(([key, value]) => `${key}: ${value}`)
        .join("\n"),
      sections: [],
    });
  }

  return {
    prefix: "company",
    content: null,
    sections,
  };
}

function buildConversationsSection(
  conversations: PlanhatConversation[]
): DustSection {
  const sections = conversations.map((conv, index) => ({
    prefix: `conversation_${index + 1}`,
    content: `Subject: ${conv.subject || "N/A"}
Type: ${conv.type || "N/A"}
Date: ${conv.date || "N/A"}
Description: ${conv.description || "N/A"}`,
    sections: [],
  }));

  return {
    prefix: "conversations",
    content: `Total conversations: ${conversations.length}`,
    sections,
  };
}

function buildEndusersSection(endusers: PlanhatEnduser[]): DustSection {
  const sections = endusers.map((user, index) => ({
    prefix: `enduser_${index + 1}`,
    content: `Name: ${user.name || "N/A"}
Email: ${user.email || "N/A"}
Role: ${user.role || "N/A"}
Status: ${user.status || "N/A"}`,
    sections: [],
  }));

  return {
    prefix: "endusers",
    content: `Total end users: ${endusers.length}`,
    sections,
  };
}

function buildNPSSection(npsData: PlanhatNPS[]): DustSection {
  const sections = npsData.map((nps, index) => ({
    prefix: `nps_${index + 1}`,
    content: `Score: ${nps.score || "N/A"}
Date: ${nps.date || "N/A"}
Comment: ${nps.comment || "N/A"}
User: ${nps.enduser?.name || "N/A"} (${nps.enduser?.email || "N/A"})`,
    sections: [],
  }));

  const avgScore =
    npsData.length > 0
      ? npsData.reduce((sum, nps) => sum + (nps.score || 0), 0) / npsData.length
      : 0;

  return {
    prefix: "nps",
    content: `Total NPS responses: ${npsData.length}
Average Score: ${avgScore.toFixed(2)}`,
    sections,
  };
}

function buildProjectsSection(projects: PlanhatProject[]): DustSection {
  const sections = projects.map((project, index) => ({
    prefix: `project_${index + 1}`,
    content: `Name: ${project.name || "N/A"}
Status: ${project.status || "N/A"}
Phase: ${project.phase || "N/A"}
Start Date: ${project.startDate || "N/A"}
End Date: ${project.endDate || "N/A"}`,
    sections: [],
  }));

  return {
    prefix: "projects",
    content: `Total projects: ${projects.length}`,
    sections,
  };
}

function buildAssetsSection(assets: PlanhatAsset[]): DustSection {
  const sections = assets.map((asset, index) => ({
    prefix: `asset_${index + 1}`,
    content: `Name: ${asset.name || "N/A"}
Type: ${asset.type || "N/A"}
Status: ${asset.status || "N/A"}
Value: ${asset.value || "N/A"}`,
    sections: [],
  }));

  return {
    prefix: "assets",
    content: `Total assets: ${assets.length}`,
    sections,
  };
}

function generateTags(company: PlanhatCompany): string[] {
  const tags: string[] = [];

  if (company.status) tags.push(`status:${company.status}`);
  if (company.health) tags.push(`health:${company.health}`);
  if (company.phase) tags.push(`phase:${company.phase}`);
  if (company.owner?.name) tags.push(`owner:${company.owner.name}`);

  return tags;
}

async function upsertToDust(document: DustDocument): Promise<void> {
  try {
    await dustLimiter.schedule(async () => {
      const {
        dataSourceId: _dsId,
        documentId,
        timestampMs,
        parents: _parents, // not supported by API
        sourceUrl,
        ...rest
      } = document;

      const body: Record<string, any> = {
        ...rest, // includes title, section, tags
        mime_type: "application/json",
        ...(timestampMs ? { timestamp: timestampMs } : {}),
      };

      if (sourceUrl) {
        body["source_url"] = sourceUrl;
      }

      const response = await dustApi.post(
        `/w/${DUST_WORKSPACE_ID}/spaces/${DUST_SPACE_ID}/data_sources/${DUST_DATASOURCE_ID}/documents/${documentId}`,
        body
      );

      if (response.status !== 200 && response.status !== 201) {
        throw new Error(
          `Failed to upsert document: ${response.status} ${response.statusText}`
        );
      }
    });
  } catch (error) {
    console.error(`Error upserting document ${document.documentId}:`, error);
    throw error;
  }
}

async function processCompany(company: PlanhatCompany): Promise<void> {
  try {
    console.log(`Processing company: ${company.name} (${company._id})`);

    // Fetch all related data
    const [conversations, endusers, nps, projects, assets] = await Promise.all([
      fetchConversations(company._id),
      fetchEndusers(company._id),
      fetchNPS(company._id),
      fetchProjects(company._id),
      fetchAssets(company._id),
    ]);

    // Build document sections
    const sections: DustSection[] = [
      buildCompanySection(company),
      buildConversationsSection(conversations),
      buildEndusersSection(endusers),
      buildNPSSection(nps),
      buildProjectsSection(projects),
      buildAssetsSection(assets),
    ];

    // Create Dust document
    const document: DustDocument = {
      dataSourceId: DUST_DATASOURCE_ID!,
      documentId: sanitizeDocumentId(company.slug || company._id),
      title: company.name,
      section: {
        prefix: "root",
        content: null,
        sections,
      },
      parents: [],
      tags: generateTags(company),
      timestampMs: company.updatedAt
        ? new Date(company.updatedAt).getTime()
        : Date.now(),
    };

    // Upsert to Dust
    await upsertToDust(document);
    console.log(`✓ Successfully processed company: ${company.name}`);
  } catch (error) {
    console.error(`✗ Error processing company ${company.name}:`, error);
  }
}

async function processCompaniesBatch(
  companies: PlanhatCompany[]
): Promise<void> {
  for (const company of companies) {
    await processCompany(company);
  }
}

// Worker thread logic
if (!isMainThread) {
  const companies: PlanhatCompany[] = workerData;
  processCompaniesBatch(companies)
    .then(() => {
      parentPort?.postMessage({ type: "done" });
    })
    .catch((error) => {
      parentPort?.postMessage({ type: "error", error: error.message });
    });
}

// Main thread logic
async function main() {
  if (!isMainThread) return;

  console.log("Starting Planhat to Dust sync...");
  console.log(`Looking back ${LOOKBACK_DAYS} days for updated companies`);
  console.log(`Using ${THREADS_NUMBER} threads for processing`);

  try {
    // Calculate lookback date
    const lookbackDate = new Date();
    lookbackDate.setDate(lookbackDate.getDate() - LOOKBACK_DAYS);

    // Fetch all companies updated since lookback date
    const companies = await fetchCompanies(lookbackDate);
    console.log(`Found ${companies.length} companies to process`);

    if (companies.length === 0) {
      console.log("No companies to process");
      return;
    }

    // Split companies into batches for parallel processing
    const batchSize = Math.ceil(companies.length / THREADS_NUMBER);
    const batches: PlanhatCompany[][] = [];

    for (let i = 0; i < companies.length; i += batchSize) {
      batches.push(companies.slice(i, i + batchSize));
    }

    // Process batches in parallel using worker threads
    const workers: Promise<void>[] = [];

    for (const batch of batches) {
      const worker = new Worker(fileURLToPath(import.meta.url), {
        workerData: batch,
      });

      const promise = new Promise<void>((resolve, reject) => {
        worker.on("message", (msg) => {
          if (msg.type === "done") {
            resolve();
          } else if (msg.type === "error") {
            reject(new Error(msg.error));
          }
        });

        worker.on("error", reject);
        worker.on("exit", (code) => {
          if (code !== 0) {
            reject(new Error(`Worker stopped with exit code ${code}`));
          }
        });
      });

      workers.push(promise);
    }

    // Wait for all workers to complete
    await Promise.all(workers);

    console.log("✅ Sync completed successfully!");
  } catch (error) {
    console.error("❌ Sync failed:", error);
    process.exit(1);
  }
}

// Run main function if this is the main thread
if (isMainThread) {
  main();
}
