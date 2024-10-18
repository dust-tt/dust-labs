import axios, { AxiosResponse } from 'axios';
import * as dotenv from 'dotenv';
import { Client } from '@notionhq/client';
import { parse } from 'csv-parse/sync';

dotenv.config();

// source
const DUST_API_KEY = process.env.DUST_API_KEY;
const DUST_WORKSPACE_ID = process.env.DUST_WORKSPACE_ID;
const DUST_USAGE_SINCE_DATE = process.env.DUST_USAGE_SINCE_DATE; // Format: YYYY-MM

// destination
const NOTION_API_KEY = process.env.NOTION_API_KEY;
const NOTION_DATABASE_ID = process.env.NOTION_DATABASE_ID;

const missingEnvVars = [
  ['DUST_API_KEY', DUST_API_KEY],
  ['DUST_WORKSPACE_ID', DUST_WORKSPACE_ID],
  ['DUST_USAGE_SINCE_DATE', DUST_USAGE_SINCE_DATE],
  ['NOTION_API_KEY', NOTION_API_KEY],
  ['NOTION_DATABASE_ID', NOTION_DATABASE_ID],
].filter(([name, value]) => !value).map(([name]) => name);

if (missingEnvVars.length > 0) {
  throw new Error(`Please provide values for the following environment variables in the .env file: ${missingEnvVars.join(', ')}`);
}

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
      const usageResponse = await dustApi.get(`/w/${DUST_WORKSPACE_ID}/workspace-usage`, {
        params: {
          // Note:
          // Ideally, what would make more sense would be to
          // have the data for the last (rolling) 30 days, just
          // like in the Dust.tt web portal.
          //
          // However, that info is not available via the public APIs.
          // The `workspace-usage` endpoint only allows for monthly granularity.
          //
          // It would be confusing for the consumer of the data if we the
          // period of time for which the usage data is reported on would
          // be changed between different runs of this script.
          //
          // Also, using the data for the previous, or the current month
          // would be misleading/non representative of the actual usage
          // of the assistants, depending on which day of the month the script
          // is run.
          //
          // The trade-off we choose is to use a random DUST_USAGE_SINCE_DATE
          // and use the current month as the end date.
          // In mid-term, I hope Dust would allow to pull the data for the last 30 days.
          start: DUST_USAGE_SINCE_DATE,
          end: new Date().toISOString().slice(0, 7), // Current month in YYYY-MM format
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
        'dust.distinctUsersReached': { number: {}, description: `Number of distinct users, since ${DUST_USAGE_SINCE_DATE}.` },
        'dust.id': { rich_text: {} },
        'dust.instructions': { rich_text: {} },
        'dust.lastVersionCreatedAt': { date: {}, description: "Last time the assistant configuration has been updated." },
        'dust.maxStepsPerRun': { number: {} },
        'dust.messages': { number: {}, description: `Number of times the assistant was used since ${DUST_USAGE_SINCE_DATE}.` },
        ...(existingDatabaseConfig.properties['dust.modelId'] ? {} : { 'dust.modelId': { select: {} } }),
        ...(existingDatabaseConfig.properties['dust.modelProviderId'] ? {} : { 'dust.modelProviderId': { select: {} } }),
        'dust.modelTemperature': { number: {} },
        'dust.pictureUrl': { url: {} },
        ...(existingDatabaseConfig.properties['dust.scope'] ? {} : { 'dust.scope': { select: {} } }),
        'dust.sId': { rich_text: {} },
        ...(existingDatabaseConfig.properties['dust.status'] ? {} : { 'dust.status': { select: {} } }),
        'dust.url': { url: {} },
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
        property: 'dust.name',
        title: { equals: assistant.name }
      }
    });

    // Map data to Notion database's properties
    const properties = {
      'dust.authors': { multi_select: assistant.authorEmails.map(author => ({ name: author })) },
      'dust.description': { rich_text: [ { text: { content: assistant.description } } ] },
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
    throw error;
  }
}

async function main() {
  console.log(`Syncing the list of Dust assistants from workspace ${DUST_WORKSPACE_ID} innto Notion database ${NOTION_DATABASE_ID}`);
  try {
    console.log(`Fetching the list of Dust assistants.`);
    const assistants = await getDustAssistants();
    console.log(`Found ${assistants.length} assistants.`);

    console.log(`Configuring the Notion database.`);
    await configureNotionDatabase();

    for (const assistant of assistants) {
      await upsertToNotion(assistant);
    }

    console.log('All assistants processed successfully.');
  } catch (error) {
    console.error('An error occurred:', error);
  }
}

main();
