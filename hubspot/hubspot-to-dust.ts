import axios from 'axios';
import * as dotenv from 'dotenv';
import Bottleneck from 'bottleneck';
import { Worker, isMainThread, parentPort, workerData } from 'worker_threads';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const HUBSPOT_ACCESS_TOKEN = process.env.HUBSPOT_ACCESS_TOKEN;
const HUBSPOT_PORTAL_ID = process.env.HUBSPOT_PORTAL_ID;
const DUST_API_KEY = process.env.DUST_API_KEY;
const DUST_WORKSPACE_ID = process.env.DUST_WORKSPACE_ID;
const DUST_DATASOURCE_ID = process.env.DUST_DATASOURCE_ID;
const DUST_VAULT_ID = process.env.DUST_VAULT_ID;

const UPDATED_SINCE_DAYS = 1; // Number of days to look back for updates
const UPDATED_SINCE = new Date(Date.now() - UPDATED_SINCE_DAYS * 24 * 60 * 60 * 1000).toISOString();
const THREADS_NUMBER = 3;

if (!HUBSPOT_ACCESS_TOKEN || !HUBSPOT_PORTAL_ID || !DUST_API_KEY || !DUST_WORKSPACE_ID || !DUST_DATASOURCE_ID) {
  throw new Error('Please provide values for HUBSPOT_ACCESS_TOKEN, HUBSPOT_PORTAL_ID, DUST_API_KEY, DUST_WORKSPACE_ID, and DUST_DATASOURCE_ID in .env file.');
}

const hubspotApi = axios.create({
  baseURL: 'https://api.hubapi.com',
  headers: {
    'Authorization': `Bearer ${HUBSPOT_ACCESS_TOKEN}`,
    'Content-Type': 'application/json'
  }
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

// Bottleneck limiter for HubSpot API
const hubspotLimiter = new Bottleneck({
  maxConcurrent: 1,
  minTime: 100 * 1.05 // 1000ms / 10 requests per second minus a 5% margin
});

// Bottleneck limiter for Dust API
const dustLimiter = new Bottleneck({
  maxConcurrent: 1,
  minTime: 500 / THREADS_NUMBER // 60000ms / 120 requests per minute
});

interface Company {
  id: string;
  properties: {
    [key: string]: string;
  };
  createdAt: string;
  updatedAt: string;
  archived: boolean;
}

interface Contact {
  id: string;
  properties: {
    [key: string]: string;
  };
  createdAt: string;
  updatedAt: string;
  archived: boolean;
}

interface Deal {
  id: string;
  properties: {
    [key: string]: string;
  };
  createdAt: string;
  updatedAt: string;
  archived: boolean;
}

interface Ticket {
  id: string;
  properties: {
    [key: string]: string;
  };
}

interface Order {
  id: string;
  properties: {
    [key: string]: string;
  };
}

interface Note {
  id: string;
  properties: {
    [key: string]: string;
  };
}

interface WorkerMessage {
  type: 'log' | 'error' | 'result';
  data: any;
}

function formatDate(dateString: string): string {
  const date = new Date(dateString);
  return date.toISOString().split('T')[0];
}

function stripHtmlTags(html: string): string {
  return html.replace(/<[^>]*>/g, '').trim();
}

async function getRecentlyUpdatedCompanyIds(): Promise<string[]> {
  let allCompanyIds: string[] = [];
  let after: string | null = null;
  const PAGE_LIMIT = 100;

  try {
    while (true) {
      const searchBody: any = {
        filterGroups: [{
          filters: [{
            propertyName: 'hs_lastmodifieddate',
            operator: 'GTE',
            value: UPDATED_SINCE
          }]
        }],
        properties: ['hs_object_id'],
        limit: PAGE_LIMIT
      };

      if (after) {
        searchBody.after = after;
      }
      console.log("call to /crm/v3/objects/companies/search")
      const response = await hubspotLimiter.schedule(() =>
        hubspotApi.post('/crm/v3/objects/companies/search', searchBody)
      );

      const companies = response.data.results;
      allCompanyIds = allCompanyIds.concat(companies.map((company: Company) => company.id));

      if (!response.data.paging?.next?.after) {
        break;
      }
      after = response.data.paging.next.after;
    }

    console.log(`Found ${allCompanyIds.length} companies with updates in the last ${UPDATED_SINCE_DAYS} day(s).`);
    return allCompanyIds;

  } catch (error) {
    console.error('Error fetching recently updated company IDs:', error);
    return [];
  }
}


async function getCompanyDetails(companyId: string): Promise<Company | null> {
  try {
    const response = await hubspotLimiter.schedule(() => hubspotApi.get(`/crm/v3/objects/companies/${companyId}`, {
      params: {
        properties: [
          'name', 'industry', 'annualrevenue', 'numberofemployees', 'phone', 'website', 'description',
          'hs_lead_status', 'createdate', 'hs_lastmodifieddate', 'lifecyclestage', 'hubspot_owner_id',
          'type', 'city', 'state', 'country', 'zip', 'address', 'facebook_company_page', 'linkedin_company_page',
          'twitterhandle', 'hs_analytics_source', 'notes_last_updated', 'hs_pipeline'
        ]
      }
    }));
    return response.data;
  } catch (error) {
    console.error(`Error fetching details for company ${companyId}:`, error);
    return null;
  }
}

async function getAssociatedContacts(companyId: string): Promise<Contact[]> {
  try {
    const response = await hubspotLimiter.schedule(() => hubspotApi.get(`/crm/v3/objects/companies/${companyId}/associations/contacts`, {
      params: {
        limit: 100
      }
    }));
    const contactIds = response.data.results.map((result: any) => result.id);

    if (contactIds.length === 0) {
      return [];
    }

    const contactsResponse = await hubspotLimiter.schedule(() => hubspotApi.post('/crm/v3/objects/contacts/batch/read', {
      properties: ['firstname', 'lastname', 'email', 'phone', 'jobtitle'],
      inputs: contactIds.map(id => ({ id }))
    }));

    return contactsResponse.data.results;
  } catch (error) {
    console.error(`Error fetching associated contacts for company ${companyId}:`, error);
    return [];
  }
}

async function getAssociatedDeals(companyId: string): Promise<Deal[]> {
  try {
    const response = await hubspotLimiter.schedule(() => hubspotApi.get(`/crm/v3/objects/companies/${companyId}/associations/deals`, {
      params: {
        limit: 100
      }
    }));
    const dealIds = response.data.results.map((result: any) => result.id);

    if (dealIds.length === 0) {
      return [];
    }

    const dealsResponse = await hubspotLimiter.schedule(() => hubspotApi.post('/crm/v3/objects/deals/batch/read', {
      properties: ['dealname', 'dealstage', 'amount', 'closedate'],
      inputs: dealIds.map(id => ({ id }))
    }));

    return dealsResponse.data.results;
  } catch (error) {
    console.error(`Error fetching associated deals for company ${companyId}:`, error);
    return [];
  }
}

async function getAssociatedTickets(companyId: string): Promise<Ticket[]> {
  try {
    const response = await hubspotLimiter.schedule(() => hubspotApi.get(`/crm/v3/objects/companies/${companyId}/associations/tickets`, {
      params: { limit: 100 }
    }));
    const ticketIds = response.data.results.map((result: any) => result.id);

    if (ticketIds.length === 0) return [];

    const ticketsResponse = await hubspotLimiter.schedule(() => hubspotApi.post('/crm/v3/objects/tickets/batch/read', {
      properties: ['subject', 'content', 'hs_pipeline_stage', 'hs_ticket_priority', 'createdate'],
      inputs: ticketIds.map(id => ({ id }))
    }));

    return ticketsResponse.data.results;
  } catch (error) {
    console.error(`Error fetching associated tickets for company ${companyId}:`, error);
    return [];
  }
}

async function getAssociatedOrders(companyId: string): Promise<Order[]> {
  try {
    const response = await hubspotLimiter.schedule(() => hubspotApi.get(`/crm/v3/objects/companies/${companyId}/associations/line_items`, {
      params: { limit: 100 }
    }));
    const orderIds = response.data.results.map((result: any) => result.id);

    if (orderIds.length === 0) return [];

    const ordersResponse = await hubspotLimiter.schedule(() => hubspotApi.post('/crm/v3/objects/line_items/batch/read', {
      properties: ['name', 'quantity', 'price', 'amount', 'createdate'],
      inputs: orderIds.map(id => ({ id }))
    }));

    return ordersResponse.data.results;
  } catch (error) {
    console.error(`Error fetching associated orders for company ${companyId}:`, error);
    return [];
  }
}

async function getNotes(companyId: string): Promise<Note[]> {
  try {
    const response = await hubspotLimiter.schedule(() => hubspotApi.get(`/crm/v3/objects/companies/${companyId}/associations/notes`, {
      params: { limit: 100 }
    }));
    const noteIds = response.data.results.map((result: any) => result.id);

    if (noteIds.length === 0) return [];

    const notesResponse = await hubspotLimiter.schedule(() => hubspotApi.post('/crm/v3/objects/notes/batch/read', {
      properties: ['hs_note_body', 'hs_createdate'],
      inputs: noteIds.map(id => ({ id }))
    }));

    return notesResponse.data.results;
  } catch (error) {
    console.error(`Error fetching notes for company ${companyId}:`, error);
    return [];
  }
}

async function upsertToDustDatasource(company: Company, contacts: Contact[], deals: Deal[], tickets: Ticket[], orders: Order[], notes: Note[]) {
  const documentId = `company-${company.id}`;
  const props = company.properties || {};

  const companyDetails = [
    `Company Name: ${props.name || 'Unknown Company'}`,
    props.industry && `Industry: ${props.industry}`,
    props.annualrevenue && `Annual Revenue: ${props.annualrevenue}`,
    props.numberofemployees && `Company Size: ${props.numberofemployees} employees`,
    props.phone && `Phone: ${props.phone}`,
    props.website && `Website: ${props.website}`,
    props.description && `Description: ${props.description}`,
    props.lifecyclestage && `Lifecycle Stage: ${props.lifecyclestage}`,
    props.hubspot_owner_id && `Owner: ${props.hubspot_owner_id}`,
    props.hs_lead_status && `Lead Status: ${props.hs_lead_status}`,
    props.type && `Type: ${props.type}`,
    props.address && `Address: ${props.address}, ${props.city || ''}, ${props.state || ''}, ${props.country || ''}, ${props.zip || ''}`,
    props.facebook_company_page && `Facebook: ${props.facebook_company_page}`,
    props.linkedin_company_page && `LinkedIn: ${props.linkedin_company_page}`,
    props.twitterhandle && `Twitter: ${props.twitterhandle}`,
    props.hs_analytics_source && `Source: ${props.hs_analytics_source}`,
    props.hs_pipeline && `Pipeline: ${props.hs_pipeline}`,
  ].filter(Boolean).join('\n');

  const contactsInfo = contacts
    .map(contact => {
      const cProps = contact.properties || {};
      const contactDetails = [
        [cProps.firstname, cProps.lastname].filter(Boolean).join(' '),
        cProps.jobtitle && `Title: ${cProps.jobtitle}`,
        cProps.email && `Email: ${cProps.email}`,
        cProps.phone && `Phone: ${cProps.phone}`,
      ].filter(Boolean);

      return contactDetails.length > 0 ? `- ${contactDetails.join(', ')}` : null;
    })
    .filter(Boolean)
    .join('\n');

  const dealsInfo = deals
    .map(deal => {
      const dProps = deal.properties || {};
      const dealDetails = [
        dProps.dealname && `${dProps.dealname}`,
        dProps.dealstage && `Stage: ${dProps.dealstage}`,
        dProps.amount && `Amount: ${dProps.amount}`,
        dProps.closedate && `Close Date: ${dProps.closedate}`,
      ].filter(Boolean);

      return dealDetails.length > 0 ? `- ${dealDetails.join(', ')}` : null;
    })
    .filter(Boolean)
    .join('\n');

  const ticketsInfo = tickets
    .map(ticket => {
      const tProps = ticket.properties || {};
      return `- ${tProps.subject || 'Untitled'}: ${tProps.hs_pipeline_stage || 'Unknown stage'}, Priority: ${tProps.hs_ticket_priority || 'Unknown'}, Created: ${tProps.createdate || 'Unknown'}`;
    })
    .join('\n');

  const ordersInfo = orders
    .map(order => {
      const oProps = order.properties || {};
      return `- ${oProps.name || 'Untitled'}: Quantity: ${oProps.quantity || '0'}, Price: ${oProps.price || '0'}, Total: ${oProps.amount || '0'}, Date: ${oProps.createdate || 'Unknown'}`;
    })
    .join('\n');

    const notesInfo = notes
    .map(note => {
      const nProps = note.properties || {};
      const formattedDate = nProps.hs_createdate ? formatDate(nProps.hs_createdate) : 'Unknown date';
      const cleanedNoteBody = stripHtmlTags(nProps.hs_note_body || 'Empty note');
      return `- ${formattedDate}: ${cleanedNoteBody}`;
    })
    .join('\n');

  const content = `
Company Summary for ${props.name || 'Unknown Company'}

Basic Company Details:
${companyDetails}

${contactsInfo ? `Key Contacts:\n${contactsInfo}` : ''}

${dealsInfo ? `Deals:\n${dealsInfo}` : ''}

${ticketsInfo ? `Tickets:\n${ticketsInfo}` : ''}

${ordersInfo ? `Orders:\n${ordersInfo}` : ''}

${notesInfo ? `Notes:\n${notesInfo}` : ''}

${props.notes_last_updated ? `Last Note Updated: ${props.notes_last_updated}` : ''}
  `.trim();

  try {
    await dustLimiter.schedule(() => dustApi.post(`/w/${DUST_WORKSPACE_ID}/vaults/${DUST_VAULT_ID}/data_sources/${DUST_DATASOURCE_ID}/documents/${documentId}`, {
      source_url: `https://app.hubspot.com/contacts/${HUBSPOT_PORTAL_ID}/company/${company.id}`,
      text: content
    }));
    console.log(`Upserted company ${company.id} to Dust datasource`);
  } catch (error) {
    console.error(`Error upserting company ${company.id} to Dust datasource:`, error);
  }
}

if (isMainThread) {
  async function main() {
    try {
      const companyIds = await getRecentlyUpdatedCompanyIds();
      console.log(`Found ${companyIds.length} companies with updates in the last ${UPDATED_SINCE_DAYS} day(s).`);

      const batchSize = Math.ceil(companyIds.length / THREADS_NUMBER);
      const batches = Array.from({ length: THREADS_NUMBER }, (_, i) =>
        companyIds.slice(i * batchSize, (i + 1) * batchSize)
      );

      const workers = batches.map((batch, index) =>
        new Worker(new URL(import.meta.url), { workerData: { batch, index } })
      );

      let processedCompanies = 0;
      await Promise.all(workers.map(worker => new Promise<void>((resolve, reject) => {
        worker.on('message', (message: WorkerMessage) => {
          if (message.type === 'log') {
            console.log(`Worker ${message.data.index}:`, message.data.message);
          } else if (message.type === 'error') {
            console.error(`Worker ${message.data.index} error:`, message.data.error);
          } else if (message.type === 'result') {
            processedCompanies += message.data;
          }
        });
        worker.on('error', reject);
        worker.on('exit', code => {
          if (code !== 0) {
            reject(new Error(`Worker stopped with exit code ${code}`));
          } else {
            resolve();
          }
        });
      })));

      console.log(`Processed ${processedCompanies} out of ${companyIds.length} companies`);
      console.log('Finished processing companies');
    } catch (error) {
      console.error('An error occurred:', error);
    }
  }

  main();
} else {
  (async () => {
    try {
      const { batch, index } = workerData as { batch: string[], index: number };
      parentPort?.postMessage({ type: 'log', data: { index, message: `Starting to process ${batch.length} companies` } });

      let processedCount = 0;
      for (const companyId of batch) {
        const company = await getCompanyDetails(companyId);
        if (company) {
          const contacts = await getAssociatedContacts(companyId);
          const deals = await getAssociatedDeals(companyId);
          const tickets = await getAssociatedTickets(companyId);
          const orders = await getAssociatedOrders(companyId);
          const notes = await getNotes(companyId);
          await upsertToDustDatasource(company, contacts, deals, tickets, orders, notes);
          processedCount++;
        }
      }

      parentPort?.postMessage({ type: 'result', data: processedCount });
      parentPort?.postMessage({ type: 'log', data: { index, message: `Finished processing ${processedCount} companies` } });
    } catch (error) {
      parentPort?.postMessage({ type: 'error', data: { index: workerData.index, error } });
    }
  })();
}

