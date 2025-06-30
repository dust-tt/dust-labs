import axios, { AxiosResponse } from "axios";
import * as dotenv from "dotenv";
import Bottleneck from "bottleneck";

dotenv.config();

const FRONT_API_TOKEN = process.env.FRONT_API_TOKEN;
const FRONT_FILTER = process.env.FRONT_FILTER;
const DUST_API_KEY = process.env.DUST_API_KEY;
const DUST_BASE_URL = "https://dust.tt/api/v1";
const DUST_WORKSPACE_ID = process.env.DUST_WORKSPACE_ID;
const DUST_SPACE_ID = process.env.DUST_SPACE_ID;
const DUST_DATASOURCE_ID = process.env.DUST_DATASOURCE_ID;

const requiredEnvVars = [
  "FRONT_API_TOKEN",
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

const DUST_TPM = parseInt(process.env.DUST_RATE_LIMIT || '120');
const FRONT_TPM = parseInt(process.env.FRONT_RATE_LIMIT || '50');
const FRONT_MAX_CONCURRENT = parseInt(process.env.FRONT_MAX_CONCURRENT || '2');
const BATCH_SIZE = parseInt(process.env.BATCH_SIZE || '50');

const frontApi = axios.create({
  baseURL: "https://api2.frontapp.com",
  headers: {
    "Content-Type": "application/json",
    "Authorization": `Bearer ${FRONT_API_TOKEN}`,
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

interface FrontInbox {
  id: string;
  name: string;
  type: string;
  is_private: boolean;
  created_at: number;
  updated_at: number;
}

interface FrontConversation {
  id: string;
  subject: string;
  status: string;
  assignee: {
    id: string;
    name: string;
    email: string;
  } | null;
  recipient: {
    handle: string;
    role: string;
  };
  tags: Array<{
    id: string;
    name: string;
  }>;
  links: {
    events: string;
    messages: string;
    comments: string;
    drafts: string;
  };
  created_at: number;
  updated_at: number;
  is_private: boolean;
  scheduled_reminders: any[];
  metadata: {
    is_first_message_outbound: boolean;
    first_message_timestamp: number;
    last_message_timestamp: number;
    customer_last_reply: number;
    customer_last_reply_at: number;
    sent_count: number;
    received_count: number;
  };
}

interface FrontConversationsResponse {
  _pagination: {
    next: string | null;
    prev: string | null;
  };
  _links: {
    self: string;
    next: string | null;
    prev: string | null;
  };
  _results: FrontConversation[];
}

interface FrontInboxesResponse {
  _pagination: {
    next: string | null;
    prev: string | null;
  };
  _links: {
    self: string;
    next: string | null;
    prev: string | null;
  };
  _results: FrontInbox[];
}

interface FrontMessage {
  id: string;
  type: string;
  created_at: number;
  updated_at: number;
  is_inbound: boolean;
  is_draft: boolean;
  is_archived: boolean;
  is_trashed: boolean;
  text: string;
  html: string;
  subject: string;
  blurb: string;
  author?: {
    id: string;
    name: string;
    email: string;
    username: string;
  };
  recipients: Array<{
    name: string;
    handle: string;
    role: string;
  }>;
  attachments: Array<{
    id: string;
    name: string;
    size: number;
    content_type: string;
    url: string;
  }>;
  metadata: {
    headers: Record<string, string>;
    is_auto_reply: boolean;
    is_bounce: boolean;
    thread_refs: string[];
  };
  conversation_id: string;
  conversation?: {
    id: string;
    subject: string;
    status: string;
    assignee: {
      id: string;
      name: string;
      email: string;
    } | null;
    recipient: {
      handle: string;
      role: string;
    };
    tags: Array<{
      id: string;
      name: string;
    }>;
    links: {
      events: string;
      messages: string;
      comments: string;
      drafts: string;
    };
    created_at: number;
    updated_at: number;
    is_private: boolean;
    scheduled_reminders: any[];
    metadata: {
      is_first_message_outbound: boolean;
      first_message_timestamp: number;
      last_message_timestamp: number;
      customer_last_reply: number;
      customer_last_reply_at: number;
      sent_count: number;
      received_count: number;
    };
  };
}

interface FrontMessagesResponse {
  _pagination: {
    next: string | null;
    prev: string | null;
  };
  _links: {
    self: string;
    next: string | null;
    prev: string | null;
  };
  _results: FrontMessage[];
}

interface RateLimitInfo {
  limit: number;
  remaining: number;
  reset: number;
}

interface Section {
  prefix?: string | null;
  content?: string | null;
  sections: Section[];
}

let currentRateLimit: RateLimitInfo | null = null;

// Function to convert URL inbox ID to API inbox ID
function convertUrlInboxIdToApiId(urlId: string): string {
  // If it already has the inb_ prefix, return as is
  if (urlId.startsWith('inb_')) {
    return urlId;
  }
  
  // Convert decimal to base-36
  const decimalId = parseInt(urlId, 10);
  if (isNaN(decimalId)) {
    throw new Error(`Invalid inbox ID format: ${urlId}. Expected decimal number or inb_ prefixed ID.`);
  }
  
  const base36Id = decimalId.toString(36);
  return `inb_${base36Id}`;
}

function updateRateLimitInfo(response: any) {
  const limit = response.headers['x-ratelimit-limit'];
  const remaining = response.headers['x-ratelimit-remaining'];
  const reset = response.headers['x-ratelimit-reset'];
  
  if (limit && remaining && reset) {
    currentRateLimit = {
      limit: parseInt(limit),
      remaining: parseInt(remaining),
      reset: parseInt(reset)
    };
    
    const resetDate = new Date(currentRateLimit.reset * 1000);
    console.log(`📊 Rate Limit: ${currentRateLimit.remaining}/${currentRateLimit.limit} requests remaining, resets at ${resetDate.toISOString()}`);
    
    // Warn if we're getting close to the limit
    if (currentRateLimit.remaining < 10) {
      console.warn(`⚠️  Rate limit warning: Only ${currentRateLimit.remaining} requests remaining!`);
    }
  }
}

function safeTimestampToISO(timestamp: number | undefined | null): string {
  if (!timestamp || timestamp <= 0 || isNaN(timestamp)) {
    return new Date().toISOString(); // Fallback to current time
  }
  
  try {
    return new Date(timestamp * 1000).toISOString();
  } catch (error) {
    console.warn(`Invalid timestamp ${timestamp}, using current time`);
    return new Date().toISOString();
  }
}

async function upsertConversationToDust(conversation: FrontConversation, messages: FrontMessage[]) {
  const conversationId = conversation.id;
  const documentId = `front-conversation-${conversationId}`;
  
  // Build conversation summary section
  const conversationSummarySection: Section = {
    prefix: "Conversation Summary",
    content: null,
    sections: [
      {
        prefix: "Details",
        content: [
          `Subject: ${conversation.subject || 'No Subject'}`,
          `Status: ${conversation.status}`,
          `Assignee: ${conversation.assignee?.name || 'Unassigned'} <${conversation.assignee?.email || 'no-email'}>`,
          `Created: ${safeTimestampToISO(conversation.created_at)}`,
          `Updated: ${safeTimestampToISO(conversation.updated_at)}`,
          `Message Count: ${messages.length}`
        ].join('\n'),
        sections: []
      }
    ]
  };

  // Build message sections with size limits
  const MAX_TOTAL_LENGTH = 2 * 1024 * 1024; // 2MB
  
  const messageSections: Section[] = [];
  let totalLength = 0;
  
  for (let i = 0; i < messages.length; i++) {
    const message = messages[i];
    
    const messageInfo = [
      `From: ${message.author?.name || 'Unknown'} <${message.author?.email || 'no-email'}>`,
      `Date: ${safeTimestampToISO(message.created_at)}`,
      `Direction: ${message.is_inbound ? 'Inbound' : 'Outbound'}`,
      `Message ID: ${message.id}`
    ].join('\n');
    
    const messageText = message.text || '';
    const messageContent = `${messageInfo}\n\n${messageText}`;
    
    // Check if adding this message would exceed the size limit
    if (totalLength + messageContent.length > MAX_TOTAL_LENGTH) {
      const remainingMessages = messages.length - i;
      messageSections.push({
        prefix: `... and ${remainingMessages} more messages (truncated due to size limit)`,
        content: null,
        sections: []
      });
      break;
    }
    
    // Determine the prefix based on message position
    let prefix: string;
    if (i === 0) {
      prefix = "Original Message";
    } else {
      prefix = `Reply ${i}`;
    }
    
    messageSections.push({
      prefix: prefix,
      content: messageContent,
      sections: []
    });
    
    totalLength += messageContent.length;
  }

  // Create the main conversation thread section
  const conversationThreadSection: Section = {
    prefix: "Conversation Thread",
    content: `Front conversation with ${messages.length} messages`,
    sections: messageSections
  };

  // Build tags array
  const tags: string[] = [
    `source:front`,
    `type:conversation`,
    `conversation_id:${conversationId}`,
    `status:${conversation.status}`,
    `message_count:${messages.length}`,
    `created_at:${safeTimestampToISO(conversation.created_at)}`,
    `updated_at:${safeTimestampToISO(conversation.updated_at)}`,
  ];

  if (conversation.assignee?.email) {
    tags.push(`assignee_email:${conversation.assignee.email}`);
  }

  if (conversation.subject) {
    tags.push(`subject:${conversation.subject}`);
  }

  // Create the main section structure
  const section: Section = {
    prefix: conversation.subject || `Conversation ${conversationId}`,
    content: `Front conversation with ${messages.length} messages`,
    sections: [conversationSummarySection, conversationThreadSection]
  };

  try {
    const response = await dustApi.post(
      `/w/${DUST_WORKSPACE_ID}/spaces/${DUST_SPACE_ID}/data_sources/${DUST_DATASOURCE_ID}/documents/${documentId}`,
      {
        section: section,
        title: `${conversation.subject || 'No Subject'}: ${messages.length} messages`,
      }
    );
    
    updateRateLimitInfo(response);
    
    console.log(`✅ Uploaded conversation ${conversationId} with ${messages.length} messages`);
  } catch (error: any) {
    console.error(`❌ Error uploading conversation ${conversationId}:`, error.message);
    if (error.response) {
      console.error("Response status:", error.response.status);
      console.error("Response data:", error.response.data);
    }
    throw error;
  }
}

async function getAllInboxes(): Promise<FrontInbox[]> {
  let allInboxes: FrontInbox[] = [];
  let nextPage: string | null = null;

  try {
    do {
      const url = nextPage || "/inboxes";
      console.log(`Fetching inboxes from: ${url}`);
      
      const response = await frontApi.get(url);
      
      // Monitor rate limits
      updateRateLimitInfo(response);
      
      const data: FrontInboxesResponse = response.data;
      
      allInboxes = allInboxes.concat(data._results);
      console.log(`Fetched ${data._results.length} inboxes (total: ${allInboxes.length})`);
      
      nextPage = data._pagination.next;
      if (nextPage) {
        const urlObj = new URL(nextPage);
        nextPage = urlObj.pathname + urlObj.search;
      }
    } while (nextPage);

    console.log(`Total inboxes fetched: ${allInboxes.length}`);
    return allInboxes;
  } catch (error: any) {
    console.error("Error fetching inboxes from Front:", error.message);
    if (error.response) {
      console.error("Response status:", error.response.status);
      console.error("Response data:", error.response.data);
    }
    throw error;
  }
}

async function getConversationsForInbox(inboxId: string, nextPageUrl: string | null = null, limit: number = 100): Promise<{ conversations: FrontConversation[], nextPage: string | null }> {
  const makeRequest = async (
    url: string,
    retryCount = 0
  ): Promise<AxiosResponse<FrontConversationsResponse>> => {
    try {
      const params: Record<string, any> = {
        limit,
      };

      console.log(`🌐 Making request to: ${url} with params:`, params);
      const response = await frontApi.get(url, { params });
      
      // Monitor rate limits
      updateRateLimitInfo(response);
      
      // Debug: Show what we actually got back
      console.log(`📊 Response: ${response.data._results.length} conversations returned`);
      if (response.data._results.length > 0) {
        const firstConv = response.data._results[0];
        console.log(`📋 Sample conversation: ${firstConv.subject} (status: ${firstConv.status}, created: ${new Date(firstConv.created_at * 1000).toISOString()})`);
      }
      
      return response;
    } catch (error: any) {
      if (error.response?.status === 429 && retryCount < 3) {
        console.log(`Rate limited, retrying in ${Math.pow(2, retryCount)} seconds...`);
        await new Promise(resolve => setTimeout(resolve, Math.pow(2, retryCount) * 1000));
        return makeRequest(url, retryCount + 1);
      }
      throw error;
    }
  };

  try {
    const url = nextPageUrl || `/inboxes/${inboxId}/conversations`;
    console.log(`Fetching conversations from: ${url}`);
    
    const response = await makeRequest(url);
    const data = response.data;
    
    console.log(`Fetched ${data._results.length} conversations for inbox ${inboxId}`);
    
    return {
      conversations: data._results,
      nextPage: data._pagination.next
    };
  } catch (error: any) {
    console.error(`Error fetching conversations for inbox ${inboxId}:`, error.message);
    if (error.response) {
      console.error("Response status:", error.response.status);
      console.error("Response data:", error.response.data);
    }
    throw error;
  }
}

async function getMessagesForConversation(conversationId: string, nextPageUrl: string | null = null, limit: number = 100): Promise<{ messages: FrontMessage[], nextPage: string | null }> {
  const makeRequest = async (
    url: string,
    retryCount = 0
  ): Promise<AxiosResponse<FrontMessagesResponse>> => {
    try {
      const params: Record<string, any> = {
        limit,
      };

      const response = await frontApi.get(url, { params });
      
      // Monitor rate limits
      updateRateLimitInfo(response);
      
      return response;
    } catch (error: any) {
      if (error.response?.status === 429 && retryCount < 3) {
        console.log(`Rate limited, retrying in ${Math.pow(2, retryCount)} seconds...`);
        await new Promise(resolve => setTimeout(resolve, Math.pow(2, retryCount) * 1000));
        return makeRequest(url, retryCount + 1);
      }
      throw error;
    }
  };

  try {
    const url = nextPageUrl || `/conversations/${conversationId}/messages`;
    console.log(`Fetching messages from: ${url}`);
    
    const response = await makeRequest(url);
    const data = response.data;
    
    console.log(`Fetched ${data._results.length} messages for conversation ${conversationId}`);
    
    return {
      messages: data._results,
      nextPage: data._pagination.next
    };
  } catch (error: any) {
    console.error(`Error fetching messages for conversation ${conversationId}:`, error.message);
    if (error.response) {
      console.error("Response status:", error.response.status);
      console.error("Response data:", error.response.data);
    }
    throw error;
  }
}

async function main() {
  console.log("🚀 Starting Front to Dust import...");
  if (FRONT_FILTER) {
    console.log(`📧 Inbox filter: ${FRONT_FILTER}`);
  } else {
    console.log(`📧 No inbox filter specified - processing all inboxes`);
  }
  console.log(`🎯 Target datasource: ${DUST_DATASOURCE_ID}`);

  try {
    // Create rate limiters
    const dustLimiter = new Bottleneck({
      maxConcurrent: 1,
      minTime: 60000 / DUST_TPM, // Convert requests per minute to milliseconds between requests
    });

    const frontLimiter = new Bottleneck({
      maxConcurrent: FRONT_MAX_CONCURRENT,
      minTime: 60000 / FRONT_TPM, // Convert requests per minute to milliseconds between requests
    });

    // Get target inbox ID from filter
    let targetInboxId: string | null = null;
    if (FRONT_FILTER) {
      try {
        const searchParams = new URLSearchParams(FRONT_FILTER);
        if (searchParams.has('inbox')) {
          const rawInboxId = searchParams.get('inbox');
          if (rawInboxId) {
            try {
              targetInboxId = convertUrlInboxIdToApiId(rawInboxId);
              console.log(`🔄 Converted inbox ID: ${rawInboxId} → ${targetInboxId}`);
            } catch (conversionError: any) {
              console.error(`❌ Error converting inbox ID "${rawInboxId}":`, conversionError.message);
              process.exit(1);
            }
          }
        }
      } catch (parseError) {
        console.warn(`Warning: Could not parse inbox filter from "${FRONT_FILTER}"`);
      }
    }

    // Get all inboxes
    console.log("📥 Fetching inboxes from Front...");
    const allInboxes = await frontLimiter.schedule(() => getAllInboxes());
    
    // Filter inboxes if specified
    let inboxesToProcess: FrontInbox[] = [];
    if (targetInboxId) {
      const targetInbox = allInboxes.find(inbox => inbox.id === targetInboxId);
      if (targetInbox) {
        inboxesToProcess = [targetInbox];
        console.log(`🎯 Processing specific inbox: ${targetInbox.name} (${targetInbox.id})`);
      } else {
        console.warn(`⚠️  Inbox with ID "${targetInboxId}" not found. Available inboxes:`);
        allInboxes.forEach(inbox => console.log(`  - ${inbox.name} (${inbox.id})`));
        return;
      }
    } else {
      inboxesToProcess = allInboxes;
      console.log(`📬 Processing all ${allInboxes.length} inboxes`);
    }

    // Process messages incrementally
    console.log("📥 Starting incremental message processing...");
    
    let totalProcessed = 0;
    let totalSuccess = 0;
    let totalErrors = 0;

    // Process each inbox
    for (const inbox of inboxesToProcess) {
      console.log(`\n📬 Processing inbox: ${inbox.name} (${inbox.id})`);
      
      let hasMoreConversations = true;
      let nextConversationPage: string | null = null;
      let conversationCount = 0;

      // Process conversations in batches
      while (hasMoreConversations) {
        try {
          // Fetch next batch of conversations
          const { conversations, nextPage: batchNextPage } = await frontLimiter.schedule(() => 
            getConversationsForInbox(inbox.id, nextConversationPage, BATCH_SIZE)
          );

          if (conversations.length === 0) {
            console.log(`✅ No more conversations in inbox ${inbox.name}`);
            break;
          }

          console.log(`📦 Processing batch of ${conversations.length} conversations...`);

          // Process each conversation
          for (const conversation of conversations) {
            conversationCount++;
            console.log(`💬 Processing conversation ${conversationCount}: ${conversation.subject || 'No Subject'} (${conversation.id})`);
            
            let hasMoreMessages = true;
            let nextMessagePage: string | null = null;
            let allMessages: FrontMessage[] = [];

            // Collect all messages for this conversation
            while (hasMoreMessages) {
              try {
                const { messages, nextPage: messageNextPage } = await frontLimiter.schedule(() => 
                  getMessagesForConversation(conversation.id, nextMessagePage, BATCH_SIZE)
                );

                if (messages.length === 0) {
                  console.log(`✅ No more messages in conversation ${conversation.id}`);
                  break;
                }

                console.log(`📧 Collecting ${messages.length} messages from conversation ${conversation.id}`);
                allMessages = allMessages.concat(messages);

                // Check if we have more messages
                if (messages.length < BATCH_SIZE) {
                  hasMoreMessages = false;
                } else {
                  nextMessagePage = messageNextPage;
                  if (!nextMessagePage) {
                    hasMoreMessages = false;
                  }
                }

              } catch (error) {
                console.error(`❌ Error collecting messages for conversation ${conversation.id}:`, error);
                console.error("🛑 Stopping script due to error");
                process.exit(1);
              }
            }

            // Upload conversation as a single document
            try {
              await dustLimiter.schedule(() => upsertConversationToDust(conversation, allMessages));
              totalProcessed += allMessages.length;
              totalSuccess += 1; // Count as 1 successful conversation upload
              console.log(`✅ Completed conversation ${conversation.id}: ${allMessages.length} messages uploaded as single document`);
            } catch (error) {
              console.error(`❌ Error uploading conversation ${conversation.id}:`, error);
              totalErrors += 1;
              throw error; // Re-throw to stop processing
            }
          }

          // Check if we have more conversations
          if (conversations.length < BATCH_SIZE) {
            hasMoreConversations = false;
          } else {
            nextConversationPage = batchNextPage;
            if (!nextConversationPage) {
              hasMoreConversations = false;
            }
          }

          // Optional: Add a small delay between conversation batches
          await new Promise(resolve => setTimeout(resolve, 1000));

        } catch (error) {
          console.error(`❌ Error processing conversations for inbox ${inbox.id}:`, error);
          console.error("🛑 Stopping script due to error");
          process.exit(1);
        }
      }

      console.log(`✅ Completed inbox ${inbox.name}: ${conversationCount} conversations processed`);
    }

    console.log("\n✅ Import completed!");
    console.log(`📊 Final Summary: ${totalProcessed} messages processed`);
    console.log(`✅ Success: ${totalSuccess} conversations imported`);
    console.log(`❌ Errors: ${totalErrors} conversations failed`);
    
    // Rate limit summary
    if (currentRateLimit) {
      console.log(`📊 Final Rate Limit Status: ${currentRateLimit.remaining}/${currentRateLimit.limit} requests remaining`);
    }
    
    if (totalErrors > 0) {
      console.log("⚠️  Some conversations failed to import. Check the logs above for details.");
    }

  } catch (error: any) {
    console.error("❌ Import failed:", error.message);
    if (error.response) {
      console.error("Response status:", error.response.status);
      console.error("Response data:", error.response.data);
    }
    process.exit(1);
  }
}

main().catch((error) => {
  console.error("Unhandled error:", error);
  process.exit(1);
}); 