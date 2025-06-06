import axios, { AxiosResponse } from "axios";
import * as dotenv from "dotenv";
import Bottleneck from "bottleneck";

dotenv.config();

const DEFAULT_FRESHSERVICE_FILTER = "updated_since=24h";

const FRESHSERVICE_DOMAIN = process.env.FRESHSERVICE_DOMAIN;
const FRESHSERVICE_API_KEY = process.env.FRESHSERVICE_API_KEY;
const FRESHSERVICE_FILTER = process.env.FRESHSERVICE_FILTER || DEFAULT_FRESHSERVICE_FILTER;
const DUST_API_KEY = process.env.DUST_API_KEY;
const DUST_BASE_URL = process.env.DUST_BASE_URL || "https://dust.tt/api/v1";
const DUST_WORKSPACE_ID = process.env.DUST_WORKSPACE_ID;
const DUST_SPACE_ID = process.env.DUST_SPACE_ID;
const DUST_DATASOURCE_ID = process.env.DUST_DATASOURCE_ID;

// Log which Dust endpoint we're using
if (process.env.DUST_BASE_URL) {
  console.log(`Using custom Dust base URL: ${DUST_BASE_URL}`);
} else {
  console.log(`Using production Dust API: ${DUST_BASE_URL}`);
}

const requiredEnvVars = [
  "FRESHSERVICE_DOMAIN",
  "FRESHSERVICE_API_KEY",
  "DUST_API_KEY",
  "DUST_WORKSPACE_ID",
  "DUST_SPACE_ID",
  "DUST_DATASOURCE_ID",
];

const missingEnvVars = requiredEnvVars.filter(
  (varName) => !process.env[varName]
);

if (missingEnvVars.length > 0) {
  throw new Error(
    `Please provide values for the following environment variables: ${missingEnvVars.join(
      ", "
    )}`
  );
}

const DUST_RATE_LIMIT = parseInt(process.env.DUST_RATE_LIMIT || '120'); // requests per minute
const FRESHSERVICE_RATE_LIMIT = parseInt(process.env.FRESHSERVICE_RATE_LIMIT || '1000'); // requests per hour
const FRESHSERVICE_MAX_CONCURRENT = parseInt(process.env.FRESHSERVICE_MAX_CONCURRENT || '2');

const freshserviceApi = axios.create({
  baseURL: `https://${FRESHSERVICE_DOMAIN}.freshservice.com/api/v2`,
  headers: {
    "Content-Type": "application/json",
    "Authorization": `Basic ${Buffer.from(FRESHSERVICE_API_KEY + ':X').toString('base64')}`,
  },
  maxContentLength: Infinity,
  maxBodyLength: Infinity,
});

const dustApi = axios.create({
  baseURL: DUST_BASE_URL,
  headers: {
    Authorization: `Bearer ${DUST_API_KEY}`,
    "Content-Type": "application/json",
  },
  maxContentLength: Infinity,
  maxBodyLength: Infinity,
});

interface FreshserviceTicket {
  id: number;
  display_id: number;
  subject: string;
  description: string;
  description_html: string;
  status: number;
  priority: number;
  source: number;
  ticket_type: string;
  category: string;
  sub_category: string;
  item_category: string;
  created_at: string;
  updated_at: string;
  due_by: string;
  fr_due_by: string;
  is_escalated: boolean;
  fr_escalated: boolean;
  spam: boolean;
  deleted: boolean;
  urgent: boolean;
  requester_id: number;
  responder_id: number | null;
  group_id: number | null;
  email_config_id: number | null;
  to_email: string | null;
  cc_email: {
    cc_emails: string[];
    fwd_emails: string[];
    reply_cc: string[];
    tkt_cc: string[];
  };
  assoc_problem_id: number | null;
  assoc_change_id: number | null;
  assoc_change_cause_id: number | null;
  assoc_asset_id: number | null;
  attachments: Array<{
    id: number;
    content_file_name: string;
    content_file_size: number;
    content_content_type: string;
    attachment_url: string;
    created_at: string;
    updated_at: string;
  }>;
  tags: Array<{
    name: string;
  }>;
  custom_field: Record<string, any>;
  status_name: string;
  priority_name: string;
  source_name: string;
  requester_name: string;
  responder_name: string;
  to_emails: string[] | null;
  department_name: string | null;
  notes: Array<{
    id: number;
    user_id: number;
    body: string;
    body_html: string;
    private: boolean;
    incoming: boolean;
    source: number;
    created_at: string;
    updated_at: string;
    deleted: boolean;
    attachments: Array<{
      id: number;
      content_file_name: string;
      content_file_size: number;
      content_content_type: string;
      attachment_url: string;
      created_at: string;
      updated_at: string;
    }>;
  }>;
  requester: {
    id: number;
    name: string;
    email: string;
    phone: string | null;
    mobile: string | null;
    department_names: string[];
    job_title: string | null;
    language: string;
    time_zone: string;
    created_at: string;
    updated_at: string;
  };
  agent?: {
    id: number;
    name: string;
    email: string;
    phone: string | null;
    mobile: string | null;
    job_title: string | null;
    language: string;
    time_zone: string;
    created_at: string;
    updated_at: string;
  };
}

interface FreshserviceTicketsResponse {
  tickets: FreshserviceTicket[];
  meta: {
    total: number;
    page: number;
    per_page: number;
    total_pages: number;
  };
}

async function getRecentTickets(): Promise<FreshserviceTicket[]> {
  let allTickets: FreshserviceTicket[] = [];
  let page = 1;
  const per_page = 30; // Freshservice default pagination

  const makeRequest = async (
    retryCount = 0
  ): Promise<AxiosResponse<FreshserviceTicketsResponse>> => {
    try {
      const params: Record<string, any> = {
        page,
        per_page,
      };

      // Only add updated_since if specified
      if (FRESHSERVICE_FILTER.includes("updated_since")) {
        const filterMatch = FRESHSERVICE_FILTER.match(/updated_since=(.+)/);
        if (filterMatch) {
          const timeValue = filterMatch[1];
          let updatedSince = "";
          
          if (timeValue === "24h") {
            const yesterday = new Date();
            yesterday.setDate(yesterday.getDate() - 1);
            updatedSince = yesterday.toISOString();
          } else if (timeValue === "1h") {
            const oneHourAgo = new Date();
            oneHourAgo.setHours(oneHourAgo.getHours() - 1);
            updatedSince = oneHourAgo.toISOString();
          } else if (timeValue === "7d") {
            const weekAgo = new Date();
            weekAgo.setDate(weekAgo.getDate() - 7);
            updatedSince = weekAgo.toISOString();
          } else {
            // Assume it's already an ISO date
            updatedSince = timeValue;
          }
          
          params.updated_since = updatedSince;
        }
      }

      console.log('Making request with params:', params);
      const response = await freshserviceApi.get("/tickets", { params });
      console.log('API Response structure:', {
        hasTickets: !!response.data.tickets,
        ticketsLength: response.data.tickets?.length,
        firstTicket: response.data.tickets?.[0] ? {
          id: response.data.tickets[0].id,
          display_id: response.data.tickets[0].display_id,
          subject: response.data.tickets[0].subject
        } : null
      });
      return response;
    } catch (error) {
      if (axios.isAxiosError(error) && error.response) {
        if (error.response.status === 429 && retryCount < 3) {
          const retryAfter = parseInt(
            error.response.headers["retry-after"] || "60",
            10
          );
          console.log(`Rate limited. Retrying after ${retryAfter} seconds...`);
          await new Promise((resolve) =>
            setTimeout(resolve, retryAfter * 1000)
          );
          return makeRequest(retryCount + 1);
        }
      }
      throw error;
    }
  };

  do {
    try {
      const response = await makeRequest();
      const tickets = response.data.tickets || [];
      
      if (Array.isArray(tickets)) {
        allTickets = allTickets.concat(tickets);
        
        // If we got less than per_page tickets, we've reached the end
        if (tickets.length < per_page) {
          break;
        }
        
        page++;
      } else {
        console.error("Unexpected response format:", response.data);
        break;
      }

      console.log(`Retrieved ${allTickets.length} valid tickets so far`);
    } catch (error) {
      if (axios.isAxiosError(error)) {
        console.error("Error fetching Freshservice tickets:");
        console.error("Status:", error.response?.status);
        console.error("Data:", JSON.stringify(error.response?.data, null, 2));
        console.error("Config:", JSON.stringify(error.config, null, 2));
      } else {
        console.error("Unexpected error:", error);
      }
      break;
    }
  } while (true);

  console.log(`Final total: ${allTickets.length} valid tickets retrieved`);
  return allTickets;
}

function formatDescription(description: string): string {
  if (!description) return "";
  
  // Remove HTML tags if description_html is used
  return description.replace(/<[^>]*>/g, "").trim();
}

function formatNotes(notes: FreshserviceTicket["notes"]): string {
  if (!notes || notes.length === 0) return "";
  
  return notes
    .filter(note => !note.deleted)
    .map(
      (note) => `
[${note.created_at}] ${note.private ? "Private" : "Public"} Note (${note.incoming ? "Incoming" : "Outgoing"}):
${formatDescription(note.body)}
${note.attachments.length > 0 ? `Attachments: ${note.attachments.map(a => a.content_file_name).join(", ")}` : ""}
`
    )
    .join("\n");
}

function formatTags(tags: Array<{ name: string }>): string {
  if (!tags || tags.length === 0) return "";
  return tags.map(tag => tag.name).join(", ");
}

function formatCCEmails(cc_email: FreshserviceTicket["cc_email"]): string {
  if (!cc_email) return "";
  
  const allEmails = [
    ...cc_email.cc_emails,
    ...cc_email.fwd_emails,
    ...cc_email.reply_cc,
    ...cc_email.tkt_cc
  ];
  
  return allEmails.length > 0 ? allEmails.join(", ") : "";
}

function formatCustomFields(custom_field: Record<string, any>): string {
  if (!custom_field || Object.keys(custom_field).length === 0) return "";
  
  return Object.entries(custom_field)
    .map(([key, value]) => `${key}: ${value}`)
    .join("\n");
}

async function upsertToDustDatasource(ticket: FreshserviceTicket) {
  if (!ticket.id) {
    console.error('Received ticket without id:', JSON.stringify(ticket, null, 2));
    throw new Error('Ticket missing id');
  }

  const documentId = `ticket-${ticket.id}`;
  console.log(`Processing ticket ID: ${ticket.id}`);
  
  const content = `
Ticket ID: ${ticket.id}
${ticket.display_id ? `Display ID: ${ticket.display_id}` : ''}
Subject: ${ticket.subject}
Description:
${formatDescription(ticket.description)}

Status: ${ticket.status_name || 'Unknown'} (${ticket.status})
Priority: ${ticket.priority_name || 'Unknown'} (${ticket.priority})
Source: ${ticket.source_name || 'Unknown'} (${ticket.source})
Type: ${ticket.ticket_type || "Not specified"}
Category: ${ticket.category || "Not specified"}
Sub Category: ${ticket.sub_category || "Not specified"}
Item Category: ${ticket.item_category || "Not specified"}

Requester: ${ticket.requester_name || 'Unknown'} (ID: ${ticket.requester_id})
${ticket.requester ? `Email: ${ticket.requester.email}` : ""}
${ticket.requester ? `Phone: ${ticket.requester.phone || "Not provided"}` : ""}
${ticket.requester ? `Department: ${ticket.requester.department_names?.join(", ") || "Not specified"}` : ""}
${ticket.requester ? `Job Title: ${ticket.requester.job_title || "Not specified"}` : ""}

Assigned Agent: ${ticket.responder_name || "Unassigned"} (ID: ${ticket.responder_id || "N/A"})
${ticket.agent ? `Agent Email: ${ticket.agent.email}` : ""}
${ticket.agent ? `Agent Phone: ${ticket.agent.phone || "Not provided"}` : ""}

Group ID: ${ticket.group_id || "Not assigned"}
Department: ${ticket.department_name || "Not specified"}

Created: ${ticket.created_at}
Updated: ${ticket.updated_at}
Due By: ${ticket.due_by || "Not set"}
First Response Due: ${ticket.fr_due_by || "Not set"}

Escalation Status:
  Is Escalated: ${ticket.is_escalated ? "Yes" : "No"}
  First Response Escalated: ${ticket.fr_escalated ? "Yes" : "No"}

Email Configuration:
  Email Config ID: ${ticket.email_config_id || "Not specified"}
  To Email: ${ticket.to_email || "Not specified"}
  CC Emails: ${formatCCEmails(ticket.cc_email || { cc_emails: [], fwd_emails: [], reply_cc: [], tkt_cc: [] })}

Flags:
  Spam: ${ticket.spam ? "Yes" : "No"}
  Deleted: ${ticket.deleted ? "Yes" : "No"}
  Urgent: ${ticket.urgent ? "Yes" : "No"}

Associated Items:
  Problem ID: ${ticket.assoc_problem_id || "None"}
  Change ID: ${ticket.assoc_change_id || "None"}
  Change Cause ID: ${ticket.assoc_change_cause_id || "None"}
  Asset ID: ${ticket.assoc_asset_id || "None"}

Tags: ${formatTags(ticket.tags || [])}

Attachments: ${(ticket.attachments || []).length > 0 ? ticket.attachments.map(a => `${a.content_file_name} (${a.content_file_size} bytes)`).join(", ") : "None"}

Custom Fields:
${formatCustomFields(ticket.custom_field || {})}

Notes and Conversations:
${formatNotes(ticket.notes || [])}
  `.trim();

  try {
    await dustApi.post(
      `/w/${DUST_WORKSPACE_ID}/spaces/${DUST_SPACE_ID}/data_sources/${DUST_DATASOURCE_ID}/documents/${documentId}`,
      {
        text: content,
      }
    );
    console.log(`Upserted ticket ID: ${ticket.id} to Dust datasource`);
  } catch (error) {
    console.error(
      `Error upserting ticket ID: ${ticket.id} to Dust datasource:`,
      error
    );
    throw error;
  }
}

async function main() {
  try {
    // Create a limiter for Freshservice API calls
    const freshserviceLimiter = new Bottleneck({
      maxConcurrent: FRESHSERVICE_MAX_CONCURRENT,
      minTime: (60 * 60 * 1000) / FRESHSERVICE_RATE_LIMIT, // Convert hourly limit to per-request delay
    });

    // Wrap the getRecentTickets function with the limiter
    const limitedGetRecentTickets = freshserviceLimiter.wrap(getRecentTickets);
    const recentTickets = await limitedGetRecentTickets();
    
    console.log(
      `Found ${recentTickets.length} tickets matching the filter criteria.`
    );

    // Create a limiter for Dust API calls
    const dustLimiter = new Bottleneck({
      maxConcurrent: 1,
      minTime: (60 * 1000) / DUST_RATE_LIMIT,
    });

    const tasks = recentTickets.map((ticket) =>
      dustLimiter.schedule(() => upsertToDustDatasource(ticket))
    );
    await Promise.all(tasks);
    console.log("All tickets processed successfully.");
  } catch (error) {
    console.error("An error occurred:", error);
    process.exit(1);
  }
}

main(); 