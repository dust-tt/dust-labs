import axios, { AxiosResponse } from 'axios';
import * as dotenv from 'dotenv';
import Bottleneck from 'bottleneck';

dotenv.config();

const ZENDESK_SUBDOMAIN = process.env.ZENDESK_SUBDOMAIN;
const ZENDESK_EMAIL = process.env.ZENDESK_EMAIL;
const ZENDESK_API_TOKEN = process.env.ZENDESK_API_TOKEN;
const DUST_API_KEY = process.env.DUST_API_KEY;
const DUST_WORKSPACE_ID = process.env.DUST_WORKSPACE_ID;
const DUST_VAULT_ID = process.env.DUST_VAULT_ID;
const DUST_DATASOURCE_ID = process.env.DUST_DATASOURCE_ID;

const missingEnvVars = [
  ['ZENDESK_SUBDOMAIN', ZENDESK_SUBDOMAIN],
  ['ZENDESK_EMAIL', ZENDESK_EMAIL],
  ['ZENDESK_API_TOKEN', ZENDESK_API_TOKEN],
  ['DUST_API_KEY', DUST_API_KEY],
  ['DUST_WORKSPACE_ID', DUST_WORKSPACE_ID],
  ['DUST_VAULT_ID', DUST_VAULT_ID],
  ['DUST_DATASOURCE_ID', DUST_DATASOURCE_ID]
].filter(([name, value]) => !value).map(([name]) => name);

if (missingEnvVars.length > 0) {
  throw new Error(`Please provide values for the following environment variables in the .env file: ${missingEnvVars.join(', ')}`);
}

const zendeskApi = axios.create({
  baseURL: `https://${ZENDESK_SUBDOMAIN}.zendesk.com/api/v2`,
  auth: {
    username: `${ZENDESK_EMAIL}/token`,
    password: ZENDESK_API_TOKEN as string
  },
  headers: {
    'Content-Type': 'application/json'
  },
  maxContentLength: Infinity,
  maxBodyLength: Infinity
});

zendeskApi.interceptors.response.use(
  (response) => {
    const rateLimit = response.headers['x-rate-limit'];
    const rateLimitRemaining = response.headers['x-rate-limit-remaining'];
    console.log(`Endpoint: ${response.config.url}, Rate Limit: ${rateLimit}, Remaining: ${rateLimitRemaining}`);
    return response;
  },
  (error) => {
    if (error.response && error.response.status === 429) {
      console.error(`Endpoint: ${error.config.url}, Rate limit exceeded. Please wait before making more requests.`);
      console.log(`Rate Limit: ${error.response.headers['x-rate-limit']}, Remaining: ${error.response.headers['x-rate-limit-remaining']}`);
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

interface Article {
  id: number;
  title: string;
  body: string;
  created_at: string;
  updated_at: string;
  author_id: number;
  section_id: number;
  url: string;
  html_url: string;
}

interface Section {
  id: number;
  name: string;
}

interface User {
  id: number;
  name: string;
}

async function getAllArticles(): Promise<Article[]> {
  let allArticles: Article[] = [];
  let currentPage = 1;

  do {
    try {
      console.log(`Fetching articles page: ${currentPage}`);
      const response: AxiosResponse<{ articles: Article[], next_page: string | null }> = await zendeskApi.get('/help_center/articles.json', {
        params: { page: currentPage }
      });

      allArticles = allArticles.concat(response.data.articles);
      console.log(`Retrieved ${response.data.articles.length} articles. Total: ${allArticles.length}`);

      // If there's no next page, it means we've fetched all articles
      if (!response.data.next_page) {
        break;
      }
      currentPage++;

    } catch (error) {
      if (axios.isAxiosError(error) && error.response?.status === 429) {
        console.error('Rate limit exceeded. Waiting before retrying...');
        await new Promise(resolve => setTimeout(resolve, 60000));
      } else {
        throw error;
      }
    }
  } while (currentPage);

  console.log(`Total articles retrieved: ${allArticles.length}`);
  return allArticles;
}

async function getSection(sectionId: number): Promise<Section | null> {
  try {
    const response: AxiosResponse<{ section: Section }> = await zendeskApi.get(`/help_center/sections/${sectionId}.json`);
    return response.data.section;
  } catch (error) {
    console.error(`Error fetching section ${sectionId}:`, error);
    return null;
  }
}

async function getUser(userId: number): Promise<User | null> {
  try {
    const response: AxiosResponse<{ user: User }> = await zendeskApi.get(`/users/${userId}.json`);
    return response.data.user;
  } catch (error) {
    console.error(`Error fetching user ${userId}:`, error);
    return null;
  }
}

async function upsertToDustDatasource(article: Article, section: Section | null, author: User | null) {
  const documentId = `article-${article.id}`;
  const content = `
Title: ${article.title}
Section: ${section ? section.name : 'Unknown'}
Author: ${author ? author.name : 'Unknown'}
Created At: ${article.created_at}
Updated At: ${article.updated_at}
Content:
${article.body}
  `.trim();

  try {
    await limitedDustApiPost(
      `/w/${DUST_WORKSPACE_ID}/vaults/${DUST_VAULT_ID}/data_sources/${DUST_DATASOURCE_ID}/documents/${documentId}`,
      {
        text: content,
        source_url: article.html_url
      }
    );
    console.log(`Upserted article ${article.id} to Dust datasource`);
  } catch (error) {
    console.error(`Error upserting article ${article.id} to Dust datasource:`, error);
  }
}

async function main() {
  try {
    const articles = await getAllArticles();
    
    for (const article of articles) {
      const section = await getSection(article.section_id);
      const author = await getUser(article.author_id);
      await upsertToDustDatasource(article, section, author);
    }

    console.log('All articles processed successfully.');
  } catch (error) {
    console.error('An error occurred:', error);
  }
}

main();
