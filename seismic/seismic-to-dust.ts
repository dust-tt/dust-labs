import axios from 'axios';
import * as dotenv from 'dotenv';
import Bottleneck from 'bottleneck';
import pdf from 'pdf-parse';

dotenv.config();

// Seismic environment variables
const SEISMIC_CLIENT_ID = process.env.SEISMIC_CLIENT_ID;
const SEISMIC_CLIENT_SECRET = process.env.SEISMIC_CLIENT_SECRET;
const SEISMIC_TENANT = process.env.SEISMIC_TENANT;
const SEISMIC_DELEGATION_USER_ID = process.env.SEISMIC_DELEGATION_USER_ID;
const SEISMIC_MAX_CONCURRENT = parseInt(process.env.SEISMIC_MAX_CONCURRENT || '2');
const SEISMIC_RATE_LIMIT_PER_MINUTE = parseInt(process.env.SEISMIC_RATE_LIMIT_PER_MINUTE || '60');
const SEISMIC_API_BASE_URL = `https://api.seismic.com`;

// Content filtering options
const CONTENT_TYPE_FILTER = process.env.CONTENT_TYPE_FILTER;
const DAYS_BACK_FILTER = process.env.DAYS_BACK_FILTER ? parseInt(process.env.DAYS_BACK_FILTER) : undefined;

// Dust API configuration
const DUST_API_KEY = process.env.DUST_API_KEY;
const DUST_WORKSPACE_ID = process.env.DUST_WORKSPACE_ID;
const DUST_SPACE_ID = process.env.DUST_SPACE_ID;
const DUST_DATASOURCE_ID = process.env.DUST_DATASOURCE_ID;
const DUST_RATE_LIMIT_PER_MINUTE = parseInt(process.env.DUST_RATE_LIMIT_PER_MINUTE || '120');

const MAX_DUST_TEXT_SIZE = 2 * 1024 * 1024; // 2MB
const SPLIT_OVERLAP = 200;

// Validate required environment variables
const requiredEnvVars = [
  'SEISMIC_CLIENT_ID', 'SEISMIC_CLIENT_SECRET', 'SEISMIC_TENANT', 'SEISMIC_DELEGATION_USER_ID',
  'DUST_API_KEY', 'DUST_WORKSPACE_ID', 'DUST_SPACE_ID', 'DUST_DATASOURCE_ID'
];

const missingVars = requiredEnvVars.filter(name => !process.env[name]);
if (missingVars.length > 0) {
  throw new Error(`Missing required environment variables: ${missingVars.join(', ')}`);
}

let seismicAccessToken: string | null = null;
let tokenExpiresAt: number = 0;

async function getSeismicAccessToken(): Promise<string> {
  if (seismicAccessToken && Date.now() < tokenExpiresAt) {
    return seismicAccessToken;
  }

  const authUrl = `https://auth.seismic.com/tenants/${SEISMIC_TENANT}/connect/token`;
  const params = new URLSearchParams({
    grant_type: 'delegation',
    client_id: SEISMIC_CLIENT_ID!,
    client_secret: SEISMIC_CLIENT_SECRET!,
    user_id: SEISMIC_DELEGATION_USER_ID!
  });

  try {
    const response = await axios.post(authUrl, params, {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });

    seismicAccessToken = response.data.access_token;
    const expiresIn = response.data.expires_in || 21600;
    tokenExpiresAt = Date.now() + (expiresIn - 300) * 1000; // Refresh 5 minutes early

    if (!seismicAccessToken) {
      throw new Error('No access token received');
    }

    return seismicAccessToken;
  } catch (error: any) {
    if (error.response?.status === 401) {
      throw new Error('Authentication failed. Please check your Seismic credentials.');
    }
    throw new Error(`Seismic authentication failed: ${error.message}`);
  }
}

const seismicApi = axios.create({
  baseURL: SEISMIC_API_BASE_URL,
  headers: { 'Content-Type': 'application/json' }
});
seismicApi.interceptors.request.use(async (config) => {
  const token = await getSeismicAccessToken();
  config.headers.Authorization = `Bearer ${token}`;
  return config;
});
seismicApi.interceptors.response.use(
  (response) => response,
  async (error) => {
    if (error.response?.status === 429) {
      await new Promise((resolve) => setTimeout(resolve, 2000));
      return seismicApi.request(error.config);
    }
    if (error.response?.status === 401) {
      seismicAccessToken = null;
      tokenExpiresAt = 0;
      const token = await getSeismicAccessToken();
      error.config.headers.Authorization = `Bearer ${token}`;
      return seismicApi.request(error.config);
    }
    return Promise.reject(error);
  }
);

const dustApi = axios.create({
  baseURL: 'https://dust.tt/api/v1',
  headers: {
    'Authorization': `Bearer ${DUST_API_KEY}`,
    'Content-Type': 'application/json'
  },
  maxContentLength: Infinity,
  maxBodyLength: Infinity
});

// Rate limiters
const seismicLimiter = new Bottleneck({
  maxConcurrent: SEISMIC_MAX_CONCURRENT,
  minTime: Math.ceil(60000 / SEISMIC_RATE_LIMIT_PER_MINUTE)
});
const dustLimiter = new Bottleneck({
  minTime: Math.ceil(60000 / DUST_RATE_LIMIT_PER_MINUTE),
  maxConcurrent: 1
});
const limitedSeismicRequest = seismicLimiter.wrap((config: any) => seismicApi.request(config));
const limitedDustRequest = dustLimiter.wrap((config: any) => dustApi.request(config));

interface SeismicContent {
  repository: "library" | "WorkSpace";
  name: string;
  teamsiteId: string;
  id: string;
  versionId: string;
  type: string;
  applicationUrls: Array<{
    name: string;
    url: string;
  }>;
  format: string;
    description?: string;
  properties?: Array<{
    name: string;
    id: string;
    values: Array<{
      id: string;
      value: string;
    }>;
  }>;
  thumbnailUrl?: string;
  downloadUrl?: string;
  createdDate?: string;
  publishDate?: string;
  modifiedDate?: string;
  majorVersion?: string;
  minorVersion?: string;
  
  // Allow additional fields from API
  [key: string]: any;
}

async function getContentDetails(contentItem?: SeismicContent): Promise<{ content: string; isMetadataOnly: boolean }> {
  let contentText: string = '';
  let metadata: any = contentItem || {};
  let isMetadataOnly = true;

  if (contentItem?.downloadUrl) {
    const response = await fetch(contentItem.downloadUrl);

    if (response.ok) {
      const contentType = response.headers.get('content-type') || '';

      if (contentType.includes('text/') || contentType.includes('json') || contentType.includes('html')) {
        contentText = await response.text();
        if (contentText && contentText.trim().length > 0) {
          isMetadataOnly = false;
        }
      } else if (contentType.includes('pdf')) {
        const blob = await response.blob();
        const arrayBuffer = await blob.arrayBuffer();
        const pdfBuffer = Buffer.from(arrayBuffer);
        
        try {
          contentText = await extractPDFText(pdfBuffer);
          if (contentText && contentText.trim().length > 0) {
            isMetadataOnly = false;
          }
        } catch (error: any) {
          console.error(`Error extracting text from PDF: ${error.message}`);
        }
      }
    } else {
      throw new Error(`Failed to download content for ${contentItem.name}: ${response.status} ${response.statusText}`);
    }
  }

  // Always try to extract metadata fallback text if we don't have meaningful content
  if (isMetadataOnly && contentItem) {
    contentText = getMetadataFallbackText(contentItem, metadata);
  }

  return { content: contentText, isMetadataOnly};
}

function getMetadataFallbackText(contentItem: SeismicContent, metadata: any): string {
  const textFields = [
    contentItem?.description,
    // Extract text from properties if available
    ...(contentItem?.properties?.map(p => p.values?.map(v => v.value).join(', ')).filter(Boolean) || []),
    // Include any additional text fields from metadata
    metadata?.description,
    metadata?.summary,
    metadata?.content,
    metadata?.text,
    metadata?.body
  ].filter(Boolean);
  
  return textFields.join('\n\n');
}

function splitContentForDust(text: string, maxSize: number, overlap: number): string[] {
  if (maxSize <= 0) {
    throw new Error('maxSize must be greater than 0');
  }
  
  // If content is smaller than maxSize, no need to split
  if (text.length <= maxSize) {
    return [text];
  }
  
  if (overlap >= maxSize) {
    overlap = Math.floor(maxSize / 2);
  }
  
  const parts: string[] = [];
  let start = 0;
  
  while (start < text.length) {
    let end = Math.min(start + maxSize, text.length);
    
    // Try to break at word boundary if not at the end
    if (end < text.length) {
      const lastSpace = text.lastIndexOf(' ', end);
      const lastNewline = text.lastIndexOf('\n', end);
      const breakPoint = Math.max(lastSpace, lastNewline);
      
      if (breakPoint > start + maxSize * 0.8) {
        end = breakPoint;
      }
    }
    
    parts.push(text.substring(start, end));
    
    // If we've reached the end, break
    if (end >= text.length) {
      break;
    }
    
    // Advance start position properly - ensure we move forward by at least (maxSize - overlap)
    start = Math.max(end - overlap, start + Math.floor(maxSize * 0.5));
  }
  
  return parts;
}

async function extractPDFText(pdfBuffer: Buffer): Promise<string> {
  // Intentionally suppress all PDF parsing logs
  const originalLog = console.log;  
  console.log = () => {};
  
  try {
    const data = await pdf(pdfBuffer);
    return data.text || '';
  } finally {
    // Restore original console methods
    console.log = originalLog;
  }
}

async function upsertToDustDatasource(seismicContent: SeismicContent, textContent: string) {
  const baseDocumentId = `seismic-${seismicContent.id}`;
    
  const metadataLines = [
    `Content Name: ${seismicContent.name}`,
    `Content ID: ${seismicContent.id}`,
    `Version ID: ${seismicContent.versionId}`,
    `Repository: ${seismicContent.repository}`,
    `TeamSite ID: ${seismicContent.teamsiteId}`,
    `Content Type: ${seismicContent.type || 'unknown'}`,
    `Format: ${seismicContent.format || 'unknown'}`,
    ...(seismicContent.majorVersion ? [`Major Version: ${seismicContent.majorVersion}`] : []),
    ...(seismicContent.minorVersion ? [`Minor Version: ${seismicContent.minorVersion}`] : []),
    ...(seismicContent.createdDate ? [`Created: ${seismicContent.createdDate}`] : []),
    ...(seismicContent.publishDate ? [`Published: ${seismicContent.publishDate}`] : []),
    ...(seismicContent.modifiedDate ? [`Modified: ${seismicContent.modifiedDate}`] : []),
    ...(seismicContent.description ? [`Description: ${seismicContent.description}`] : []),
    ...(seismicContent.properties && seismicContent.properties.length > 0 ? 
        [`Properties: ${seismicContent.properties.map((p: any) => `${p.name}: ${p.values?.map((v: any) => v.value).join(', ') || 'N/A'}`).join('; ')}`] : [])
  ].join('\n');

  const baseText = `---\n${metadataLines}\n---\n\n`;
  const finalContent = baseText + textContent;
  
  const maxSizeOfPartMeta = `\n(Part 999 of 999)`;
  const maxPartMetaOverhead = Buffer.byteLength(maxSizeOfPartMeta, 'utf8');
  const maxContentSize = MAX_DUST_TEXT_SIZE - maxPartMetaOverhead;
  
  if (maxContentSize <= 0) {
    throw new Error(`Part metadata overhead is too large for ${seismicContent.name}. Cannot split content.`);
  }
  
  const contentParts = splitContentForDust(finalContent, maxContentSize, SPLIT_OVERLAP);

  for (let i = 0; i < contentParts.length; i++) {
    const documentId = contentParts.length === 1 ? baseDocumentId : `${baseDocumentId}-part${i + 1}`;
    const partMeta = contentParts.length === 1 ? '' : `\n(Part ${i + 1} of ${contentParts.length})`;
    const textWithPart = contentParts[i] + partMeta;
    
    await limitedDustRequest({
      method: 'POST',
      url: `/w/${DUST_WORKSPACE_ID}/spaces/${DUST_SPACE_ID}/data_sources/${DUST_DATASOURCE_ID}/documents/${documentId}`,
      data: {
        text: textWithPart,
        source_url: seismicContent.applicationUrls.find(url => url)?.url,
        title: contentParts.length === 1 ? seismicContent.name : `${seismicContent.name} (Part ${i + 1} of ${contentParts.length})`,
        mime_type: 'text/plain'
      }
    });
  }
}

async function processAndSyncLibraryContents(contentType?: string): Promise<void> {
  const filterDescriptions = [
    contentType ? `type '${contentType}'` : null,
    DAYS_BACK_FILTER ? `created in the last ${DAYS_BACK_FILTER} days` : null
  ].filter(Boolean);
  
  console.log(`Processing content${filterDescriptions.length > 0 ? ` (${filterDescriptions.join(', ')})` : ''}...`);

  let continuationToken: string | null = null;
  let pageNumber = 1;
  let totalProcessed = 0;
  let totalFound = 0;
  
  // Calculate date threshold for filtering
  let createdDateThreshold: string | undefined;
  if (DAYS_BACK_FILTER) {
    const thresholdDate = new Date();
    thresholdDate.setDate(thresholdDate.getDate() - DAYS_BACK_FILTER);
    createdDateThreshold = thresholdDate.toISOString().split('.')[0] + 'Z';
  }
  
  do {
    // Build filter conditions
    const filterConditions: any[] = [];
    
    if (contentType) {
      filterConditions.push({ attribute: "format", operator: "equal", value: contentType });
    }
    
    if (createdDateThreshold) {
      filterConditions.push({ attribute: "createdDate", operator: "greaterThanOrEqual", value: createdDateThreshold });
    }
    
    const requestData = {
      options: {
        pageSize: 100,
        returnFields: [
          "repository", "name", "teamsiteId", "id", "versionId", "type", "applicationUrls", "format",
          "description", "properties", "thumbnailUrl", "downloadUrl", 
          "createdDate", "publishDate", "modifiedDate", "majorVersion", "minorVersion"
        ]
      },
      ...(filterConditions.length > 0 && {
        filter: {
          operator: "and",
          conditions: filterConditions
        }
      })
    };
    
    const url: string = continuationToken 
      ? `/search/v1/content/query?continuationToken=${encodeURIComponent(continuationToken)}`
      : '/search/v1/content/query';
    
    const searchResponse: any = await limitedSeismicRequest({
      method: 'POST',
      url,
      ...(continuationToken ? {} : { data: requestData })
    });

    const searchResults = searchResponse.data.documents || [];
    totalFound += searchResults.length;
    
    console.log(`Page ${pageNumber}: Found ${searchResults.length} items`);
    
    // Process each content item
    for (const content of searchResults) {
      try {
        console.log(`Processing: ${content.name}`);
        
        const { content: contentData, isMetadataOnly } = await getContentDetails(content);
        
        await upsertToDustDatasource(content, contentData);
        
        if (isMetadataOnly) {
          console.log(`Processed as metadata-only: ${content.name}`);
        }
        
        totalProcessed++;
      } catch (error: any) {
        console.error(`Error processing ${content.name}: ${error.message}`);
      }
    }
    
    continuationToken = searchResponse.data.continuationToken || null;
    pageNumber++;
  } while (continuationToken);
  
  console.log(`\nProcessed ${totalProcessed} items`);
}

async function main() {
  try {
    console.log('Starting Seismic to Dust synchronization...');
    
    await processAndSyncLibraryContents(CONTENT_TYPE_FILTER || undefined);
    
    console.log('Synchronization completed successfully.');
  } catch (error: any) {
    console.error('Synchronization failed:', error.message);
    process.exit(1);
  }
}

main();