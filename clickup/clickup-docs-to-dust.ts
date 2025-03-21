import axios, { AxiosResponse } from 'axios';
import * as dotenv from 'dotenv';
import Bottleneck from 'bottleneck';
import slugify from 'slugify';

dotenv.config();

// source
const CLICKUP_API_KEY = process.env.CLICKUP_API_KEY;
const CLICKUP_WORKSPACE_ID = process.env.CLICKUP_WORKSPACE_ID;
const CLICKUP_DOC_ID = process.env.CLICKUP_DOC_ID;

// destination
const DUST_API_KEY = process.env.DUST_API_KEY;
const DUST_WORKSPACE_ID = process.env.DUST_WORKSPACE_ID;
const DUST_DATASOURCE_ID = process.env.DUST_DATASOURCE_ID;
const DUST_VAULT_ID = process.env.DUST_VAULT_ID;

const missingEnvVars = [
  ['CLICKUP_API_KEY', CLICKUP_API_KEY],
  ['CLICKUP_WORKSPACE_ID', CLICKUP_WORKSPACE_ID],
  ['CLICKUP_DOC_ID', CLICKUP_DOC_ID],
  ['DUST_API_KEY', DUST_API_KEY],
  ['DUST_WORKSPACE_ID', DUST_WORKSPACE_ID],
  ['DUST_DATASOURCE_ID', DUST_DATASOURCE_ID],
  ['DUST_VAULT_ID', DUST_VAULT_ID]
].filter(([name, value]) => !value).map(([name]) => name);

if (missingEnvVars.length > 0) {
  throw new Error(`Please provide values for the following environment variables in the .env file: ${missingEnvVars.join(', ')}`);
}

const clickupApi = axios.create({
  baseURL: 'https://api.clickup.com/api/v3',
  headers: {
    'Authorization': CLICKUP_API_KEY,
    'Content-Type': 'application/json'
  },
  maxContentLength: Infinity,
  maxBodyLength: Infinity
});

clickupApi.interceptors.response.use(
  (response) => {
    console.log(`Endpoint: ${response.config.url}, Status: ${response.status}`);
    return response;
  },
  (error) => {
    if (error.response && error.response.status === 429) {
      console.error(`Endpoint: ${error.config.url}, Rate limit exceeded. Please wait before making more requests.`);
    }
    return Promise.reject(error);
  }
);

// Create a Bottleneck limiter for Dust API
const dustLimiter = new Bottleneck({
  minTime: 500, // 500ms between requests
  maxConcurrent: 1, // Only 1 request at a time
});

const dustApi = axios.create({
  baseURL: 'https://dust.tt/api/v1',
  headers: {
    'Authorization': `Bearer ${DUST_API_KEY}`,
    'Content-Type': 'application/json'
  },
  maxContentLength: Infinity,
  maxBodyLength: Infinity
});

// Wrap dustApi.post with the limiter
const limitedDustApiPost = dustLimiter.wrap(
  (url: string, data: any, config?: any) => dustApi.post(url, data, config)
);

interface ClickUpPage {
  workspace_id: string;
  doc_id: string;
  id: string;
  archived: boolean;
  name: string;
  content: string;
  parent_id: string | null;
  pages: ClickUpPage[];
  date_created: number;
  date_updated: number;
}

function generateSlug(name: string): string {
  return slugify(name, {
    lower: true,
    strict: true,
    trim: true
  });
}

async function getClickUpPages(docId: string): Promise<ClickUpPage[]> {
  try {
    const response: AxiosResponse<ClickUpPage[]> = await clickupApi.get(
      `/workspaces/${CLICKUP_WORKSPACE_ID}/docs/${docId}/pages`,
      {
        params: {
          max_page_depth: -1,
          content_format: 'text/md'
        }
      }
    );
    console.log(`Retrieved ${response.data.length} pages from ClickUp`);
    return response.data;
  } catch (error) {
    console.error('Error fetching ClickUp pages:', error);
    throw error;
  }
}

async function upsertToDustDatasource(page: ClickUpPage) {
  const slug = generateSlug(page.name);
  const documentId = `${page.id}-${slug}`;
  const createdDate = new Date(page.date_created).toISOString();
  const updatedDate = new Date(page.date_updated).toISOString();

  const content = `
Title: ${page.name}
Created At: ${createdDate}
Updated At: ${updatedDate}
Content:
${page.content}
  `.trim();

  try {
    await limitedDustApiPost(
      `/w/${DUST_WORKSPACE_ID}/vaults/${DUST_VAULT_ID}/data_sources/${DUST_DATASOURCE_ID}/documents/${documentId}`,
      {
        text: content,
        source_url: `https://app.clickup.com/${CLICKUP_WORKSPACE_ID}/v/dc/${CLICKUP_DOC_ID}/${page.id}`
      }
    );
    console.log(`Upserted page '${documentId}' https://app.clickup.com/${CLICKUP_WORKSPACE_ID}/v/dc/${CLICKUP_DOC_ID}/${page.id} to Dust datasource`);
  } catch (error) {
    console.error(`Error upserting page '${documentId}') to Dust datasource:`, error);
  }
}

async function processPages(pages: ClickUpPage[]) {
  for (const page of pages) {

    // skip empty pages
    if (page.content && page.content.trim() !== '') {
      if (!page.archived) {
        await upsertToDustDatasource(page);
      } else {
        console.log(`Skipping archived page: ${page.name}`);
      }
    }

    if (page.pages && page.pages.length > 0) {
      await processPages(page.pages);
    }
  }
}

async function main() {
  try {
    const pages = await getClickUpPages(CLICKUP_DOC_ID!);
    await processPages(pages);
    console.log('All pages processed successfully.');
  } catch (error) {
    console.error('An error occurred:', error);
  }
}

main();
