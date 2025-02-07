import axios, { AxiosResponse } from 'axios';
import * as dotenv from 'dotenv';
import { Client, iteratePaginatedAPI } from '@notionhq/client';
import { parse } from 'csv-parse/sync';

dotenv.config();

// source
const DUST_API_KEY = process.env.DUST_API_KEY;
const DUST_WORKSPACE_ID = process.env.DUST_WORKSPACE_ID;

// destination
const NOTION_API_KEY = process.env.NOTION_API_KEY;
const NOTION_DATABASE_ID = process.env.NOTION_DATABASE_ID;

// configurable behavior
const NOTION_SOFT_DELETE = process.env.NOTION_SOFT_DELETE === 'true' || false;

const missingEnvVars = [
  ['DUST_API_KEY', DUST_API_KEY],
  ['DUST_WORKSPACE_ID', DUST_WORKSPACE_ID],
  ['NOTION_API_KEY', NOTION_API_KEY],
  ['NOTION_DATABASE_ID', NOTION_DATABASE_ID],
].filter(([name, value]) => !value).map(([name]) => name);

if (missingEnvVars.length > 0) {
  throw new Error(`Please provide values for the following environment variables in the .env file: ${missingEnvVars.join(', ')}`);
}

// Last 30 (rolling) days
const usage_start = new Date(Date.now() - 31 * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
const usage_end = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString().slice(0, 10);

interface DustAssistant {
  id: number;
  sId: string;
  scope: string;
  name: string;
  pictureUrl: string;
  description: string;
  instructions: string;
  model: {
    providerId: string;
    modelId: string;
    temperature: number;
  };
  status: string;
  maxStepsPerRun: number;
  versionCreatedAt: string;
  visualizationEnabled: boolean;
  templateId: string;
}

const dustApi = axios.create({
  baseURL: 'https://dust.tt/api/v1',
  headers: {
    'Authorization': `Bearer ${DUST_API_KEY}`,
    'Content-Type': 'application/json'
  },
  maxContentLength: Infinity,
  maxBodyLength: Infinity,
});

async function getDustAssistants(): Promise<DustAssistant[]> {
  try {

    // Fetch list of assistants
    const response: AxiosResponse<{
      agentConfigurations: DustAssistant[];
    }> = await dustApi.get(`/w/${DUST_WORKSPACE_ID}/assistant/agent_configurations`);

    let assistants = response.data.agentConfigurations.map((assistant: DustAssistant) => ({
      id: assistant.id,
      sId: assistant.sId,
      scope: assistant.scope,
      name: assistant.name,
      pictureUrl: assistant.pictureUrl,
      description: assistant.description,
      instructions: assistant.instructions,
      model: assistant.model,
      status: assistant.status,
      maxStepsPerRun: assistant.maxStepsPerRun,
      versionCreatedAt: assistant.versionCreatedAt,
      visualizationEnabled: assistant.visualizationEnabled,
      templateId: assistant.templateId,
    }));

    // Fetch assistants' usage data
    try {
      console.log(`Fetching assistants' usage data between ${usage_start} and ${usage_end}`);
      const usageResponse = await dustApi.get(`/w/${DUST_WORKSPACE_ID}/workspace-usage`, {
        params: {
          start: usage_start,
          end: usage_end,
          mode: 'range',
          table: 'assistants'
        }
      });

      // Parse CSV data using csv-parse
      const usageData = parse(usageResponse.data, {
        columns: true,
        skip_empty_lines: true
      });

      // Enrich assistants with usage data
      assistants = assistants.map(assistant => {
        const usage = usageData.find(u => u.name === assistant.name);
        if (!usage) {
          console.warn(`Warning: Usage data not found for assistant "${assistant.name}"`);
        }
        return {
          ...assistant,
          authorEmails: usage ? JSON.parse(usage.authorEmails) : [],
          messages: usage ? parseInt(usage.messages, 10) : 0,
          distinctConversations: usage ? parseInt(usage.distinctConversations, 10) : 0,
          distinctUsersReached: usage ? parseInt(usage.distinctUsersReached, 10) : 0,
        };
      });
    } catch (error) {
      console.error('Warning: usage data will not be updated. We encountered an error while fetching assistants\' usage data:', error);
    }

    return assistants as DustAssistant[];
  } catch (error) {
    if (axios.isAxiosError(error) && error.response) {
      console.error('Error fetching Dust assistants:', error.response.data);
    } else {
      console.error('Error fetching Dust assistants:', error);
    }
    throw error;
  }
}

const notion = new Client({ auth: NOTION_API_KEY });

async function configureNotionDatabase() {

  // Get existing database configuration
  let existingDatabaseConfig;
  try {
    existingDatabaseConfig = await notion.databases.retrieve({
      database_id: NOTION_DATABASE_ID ?? '',
    });
    console.log('Retrieved existing Notion database configuration');
  } catch (error) {
    console.error('Error retrieving existing Notion database configuration:', error);
  }
  try {
    const response = await notion.databases.update({
      database_id: NOTION_DATABASE_ID ?? '',
      description: [
        { text: { content: "💡 All the ", } },
        { text: { content: "dust.*", }, annotations: { "code": true, } },
        { text: { content: " fields are automatically synced from Dust.tt. Last sync: ", } },
        { text: { content: new Date().toLocaleString(), }, annotations: { code: true, } },
      ],
      properties: {
        ...(existingDatabaseConfig.properties.Name ? { Name: { name: "dust.name"} } : {}), // Rename the 'Name' property to 'dust.name'
        ...(existingDatabaseConfig.properties.Tags ? { Tags: null } : {}), // Remove the 'Tags' property if it exists in the current configuration
        ...(existingDatabaseConfig.properties['dust.authors'] ? {} : { 'dust.authors': { multi_select: {} } }),
        'dust.description': { rich_text: {} },
        'dust.distinctConversations': { number: {}, description: `Number of distinct conversations over the last 30 days.` },
        'dust.distinctUsersReached': { number: {}, description: `Number of distinct users over the last 30 days.` },
        'dust.id': { rich_text: {} },
        'dust.instructions': { rich_text: {} },
        'dust.lastVersionCreatedAt': { date: {}, description: "Last time the assistant configuration has been updated." },
        'dust.maxStepsPerRun': { number: {} },
        'dust.messages': { number: {}, description: `Number of messages the assistant has sent over the last 30 days.` },
        ...(existingDatabaseConfig.properties['dust.modelId'] ? {} : { 'dust.modelId': { select: {} } }),
        ...(existingDatabaseConfig.properties['dust.modelProviderId'] ? {} : { 'dust.modelProviderId': { select: {} } }),
        'dust.modelTemperature': { number: {} },
        'dust.pictureUrl': { url: {} },
        ...(existingDatabaseConfig.properties['dust.scope'] ? {} : { 'dust.scope': { select: {} } }),
        'dust.sId': { rich_text: {} },
        ...(existingDatabaseConfig.properties['dust.status'] ? {} : { 'dust.status': { select: {} } }),
        'dust.url': { url: {} },
        'dust.assistantDetailsUrl': { url: {} },
        'dust.visualizationEnabled': { checkbox: {} },
      }
    });
    console.log('Notion database configured successfully');
    return response;
  } catch (error) {
    console.error('Error configuring Notion database:', error);
    throw error;
  }
}

async function upsertToNotion(assistant: any) {
  try {
    // Check if the page already exists
    const existingPages = await notion.databases.query({
      database_id: NOTION_DATABASE_ID ?? '',
      filter: {
        property: 'dust.sId',
        rich_text: { equals: assistant.sId }
      }
    });

    // Map data to Notion database's properties
    const properties = {
      'dust.authors': { multi_select: assistant.authorEmails.map(author => ({ name: author })) },
      'dust.description': { rich_text: [ { text: { content: assistant.description } } ] },
      'dust.distinctConversations': { number: assistant.distinctConversations },
      'dust.distinctUsersReached': { number: assistant.distinctUsersReached },
      'dust.id': { rich_text: [ { text: { content: assistant.id.toString() } } ] },
      'dust.instructions': { rich_text: [ { text: { content: (assistant.instructions || '').substring(0, 2000) } } ] },
      'dust.lastVersionCreatedAt': { date: assistant.versionCreatedAt ? { start: assistant.versionCreatedAt } : null },
      'dust.maxStepsPerRun': { number: assistant.maxStepsPerRun },
      'dust.messages': { number: assistant.messages },
      'dust.modelId': { select: { name: assistant.model.modelId } },
      'dust.modelProviderId': { select: { name: assistant.model.providerId } },
      'dust.modelTemperature': { number: assistant.model.temperature },
      'dust.name': { title: [ { text: { content: assistant.name } } ] },
      'dust.pictureUrl': { url: assistant.pictureUrl || null },
      'dust.scope': { select: { name: assistant.scope } },
      'dust.sId': { rich_text: [ { text: { content: assistant.sId } } ] },
      'dust.status': { select: { name: assistant.status } },
      'dust.url': { url: `https://dust.tt/w/${DUST_WORKSPACE_ID}/assistant/new?assistant=${assistant.sId}` },
      'dust.assistantDetailsUrl': { url: `https://dust.tt/w/${DUST_WORKSPACE_ID}/assistant/new?assistantDetails=${assistant.sId}` },
      'dust.visualizationEnabled': { checkbox: assistant.visualizationEnabled },
    }

    let response;
    if (existingPages.results.length > 0) { // update existing entry if there is (at least) a match
      console.log(`Updating assistant '${assistant.name}' in Notion database`);
      response = await notion.pages.update({
        page_id: existingPages.results[0].id,
        properties: properties as Record<string, any>
      });
    } else { // create new entry if there is no match
      console.log(`Creating assistant '${assistant.name}' in Notion database`);
      response = await notion.pages.create({
        parent: { database_id: NOTION_DATABASE_ID ?? '' },
        properties: properties as Record<string, any>
      });
    }

    return response;
  } catch (error) {
    console.error(`Error processing assistant '${assistant.name}':`, error);
    return null;
  }
}

async function processOrphanedPages(assistants: DustAssistant[], softDelete: boolean) {
  console.log('Checking for orphaned pages in Notion database...');
  const existingAssistantSIds = new Set(assistants.map(a => a.sId));

  for await (const page of iteratePaginatedAPI(notion.databases.query, {
    database_id: NOTION_DATABASE_ID ?? '',
    filter: {
      property: 'dust.status',
      select: {
        equals: 'active'
      }
    }
  })) {
    if ('properties' in page) {
      const sIdProp = page.properties['dust.sId'];
      const nameProp = page.properties['dust.name'];

      if (sIdProp && 'rich_text' in sIdProp && nameProp && 'title' in nameProp) {
        const pageSId = sIdProp.rich_text[0]?.plain_text;
        const assistantName = nameProp.title[0]?.plain_text;
        if (pageSId && !existingAssistantSIds.has(pageSId)) {
          await notion.comments.create({
            parent: { page_id: page.id },
            rich_text: [{ text: { content: `This assistant has been deleted on ${new Date().toLocaleString()} because it no longer appears as a shared assistant in Dust.` } }]
          });
          if (softDelete) {
            console.log(`Soft-deleting assistant "${assistantName}" (${pageSId})`);
            await notion.pages.update({
              page_id: page.id,
              properties: {
                'dust.status': { select: { name: 'deleted' } },
              },
            });
          } else {
            console.log(`Deleting assistant "${assistantName}" (${pageSId})`);
            await notion.pages.update({
              page_id: page.id,
              archived: true,
            });
          }
        }
      }
    }
  }

  console.log('Finished checking for orphaned pages.');
}

async function main() {
  console.log(`Syncing the list of Dust assistants from workspace ${DUST_WORKSPACE_ID} into Notion database ${NOTION_DATABASE_ID}`);
  try {
    console.log(`Fetching the list of Dust assistants.`);
    const assistants = await getDustAssistants();
    console.log(`Found ${assistants.length} assistants.`);

    console.log(`Configuring the Notion database.`);
    await configureNotionDatabase();

    for (const assistant of assistants) {
      await upsertToNotion(assistant);
    }

    await processOrphanedPages(assistants, NOTION_SOFT_DELETE);

    console.log('All assistants processed successfully.');
  } catch (error) {
    console.error('An error occurred:', error);
  }
}

main();
