import axios from "axios";
import * as dotenv from "dotenv";
import Bottleneck from "bottleneck";
import * as talon from "talonjs";
import { Front } from "front-sdk";
import type { 
  Inbox, 
  Conversation, 
  Message,
  Conversations,
  ConversationMessages
} from "front-sdk";

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

// Time filtering configuration
const DAYS_BACK = parseInt(process.env.DAYS_BACK || '30'); // Default to 30 days back
const cutoffTimestamp = Math.floor((Date.now() - (DAYS_BACK * 24 * 60 * 60 * 1000)) / 1000);

// Message count filtering
const MIN_MESSAGE_COUNT = parseInt(process.env.MIN_MESSAGE_COUNT || '1');

const front = new Front(FRONT_API_TOKEN!);

const dustApi = axios.create({
  baseURL: DUST_BASE_URL,
  headers: {
    Authorization: `Bearer ${DUST_API_KEY}`,
    "Content-Type": "application/json",
  },
  maxContentLength: Infinity,
  maxBodyLength: Infinity,
});

interface TimelineEntry {
  timestamp: Date;
  type: "RECEIVED" | "SENT";
  sender: string;
  recipients: string[];
  subject?: string;
  content: string;
  attachments: string[];
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

function convertUrlInboxIdToApiId(urlId: string): string {
  if (urlId.startsWith('inb_')) {
    return urlId;
  }
  
  const decimalId = parseInt(urlId, 10);
  if (isNaN(decimalId)) {
    throw new Error(`Invalid inbox ID format: ${urlId}. Expected decimal number or inb_ prefixed ID.`);
  }
  
  const base36Id = decimalId.toString(36);
  return `inb_${base36Id}`;
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

function safeTimestampToDate(timestamp: number | undefined | null): Date {
  if (!timestamp || timestamp <= 0 || isNaN(timestamp)) {
    return new Date(); // Fallback to current time
  }
  
  try {
    return new Date(timestamp * 1000);
  } catch (error) {
    console.warn(`Invalid timestamp ${timestamp}, using current time`);
    return new Date();
  }
}

async function upsertConversationToDust(conversation: Conversation, messages: Message[]) {
  const conversationId = conversation.id;
  const documentId = `front-conversation-${conversationId}`;
  
  const timeline = createTimelineFromMessages(messages);
  let llmContent = timelineToLLMFormat(timeline);
  
  // Apply size limits (2MB)
  const MAX_TOTAL_LENGTH = 2 * 1024 * 1024;
  if (llmContent.length > MAX_TOTAL_LENGTH) {
    llmContent = llmContent.substring(0, MAX_TOTAL_LENGTH - 200) + 
      "\n\n[CONTENT TRUNCATED DUE TO SIZE LIMIT - SOME MESSAGES MAY BE MISSING]";
  }

  const tags: string[] = [
    `source:front`,
    `type:conversation`,
    `conversation_id:${conversationId}`,
    `status:${conversation.status}`,
    `message_count:${messages.length}`,
    `created_at:${safeTimestampToISO(conversation.created_at)}`,
  ];

  if (conversation.assignee?.email) {
    tags.push(`assignee_email:${conversation.assignee.email}`);
  }

  if (conversation.subject) {
    tags.push(`subject:${conversation.subject}`);
  }

  const section: Section = {
    prefix: conversation.subject || `Conversation ${conversationId}`,
    content: llmContent,
    sections: []
  };

  try {
    await dustApi.post(
      `/w/${DUST_WORKSPACE_ID}/spaces/${DUST_SPACE_ID}/data_sources/${DUST_DATASOURCE_ID}/documents/${documentId}`,
      {
        section: section,
        title: `${conversation.subject || 'No Subject'}: ${messages.length} messages`,
        tags: tags,
      }
    );
        
    console.log(`Uploaded conversation ${conversationId} with ${messages.length} messages`);
  } catch (error: any) {
    console.error(`Error uploading conversation ${conversationId}:`, error.message);
    throw error;
  }
}

async function getAllInboxes(): Promise<Inbox[]> {
  try {
    console.log("Fetching inboxes...");
    const inboxes: Inbox[] = [];

    let nextPageUrl: string | null = null;
    while (true) {
      const response = await front.inbox.list();
      inboxes.push(...response._results);
      nextPageUrl = response._pagination?.next ?? null;
      if (!nextPageUrl) {
        break;
      }
    }
    console.log(`Found ${inboxes.length} inboxes`);
    return inboxes;
  } catch (error: any) {
    console.error("Error fetching inboxes:", error.message);
    throw error;
  }
}

async function getConversationsForInbox(inboxId: string, nextPageUrl: string | null = null, limit: number = 100): Promise<{ conversations: Conversation[], nextPage: string | null, hasMoreRecent: boolean }> {
  const makeRequest = async (retryCount = 0): Promise<Conversations> => {
    try {
      const params: any = { 
        inbox_id: inboxId,
        limit,
        page_token: nextPageUrl ? new URL(nextPageUrl).searchParams.get('page_token') : null,
      };
      
      return await front.inbox.listConversations(params);
    } catch (error: any) {
      if (error.message?.includes('429') && retryCount < 3) {
        console.log(`Rate limited, retrying in ${Math.pow(2, retryCount)} seconds...`);
        await new Promise(resolve => setTimeout(resolve, Math.pow(2, retryCount) * 1000));
        return makeRequest(retryCount + 1);
      }
      throw error;
    }
  };

  try {
    const data = await makeRequest();
    
    let filteredConversations = data._results;
    let hasMoreRecent = true;

    const beforeFilter = filteredConversations.length;
    filteredConversations = filteredConversations.filter(conv => {
      if (!conv.created_at || conv.created_at <= 0 || isNaN(conv.created_at)) {
        console.warn(`Invalid created_at timestamp for conversation ${conv.id}: ${conv.created_at}`);
        return false;
      }
      return conv.created_at >= cutoffTimestamp;
    });
    
    const afterFilter = filteredConversations.length;
    
    if (beforeFilter > afterFilter) {
      const oldestFiltered = data._results[data._results.length - 1];
      if (oldestFiltered && oldestFiltered.created_at < cutoffTimestamp) {
        hasMoreRecent = false;
        console.log(`Found conversations older than ${DAYS_BACK} days, stopping pagination`);
      }
    }
    
    return {
      conversations: filteredConversations,
      nextPage: data._pagination?.next ?? null,
      hasMoreRecent
    };
  } catch (error: any) {
    console.error(`Error fetching conversations for inbox ${inboxId}:`, error.message);
    throw error;
  }
}

async function getMessagesForConversation(conversationId: string, nextPageUrl: string | null = null, limit: number = 100): Promise<{ messages: Message[], nextPage: string | null, hasMoreRecent: boolean }> {
  const makeRequest = async (retryCount = 0): Promise<ConversationMessages> => {
    try {
      const params: any = { 
        conversation_id: conversationId,
        limit,
        page_token: nextPageUrl ? new URL(nextPageUrl).searchParams.get('page_token') : null,
      };
      
      return await front.conversation.listMessages(params);
    } catch (error: any) {
      if (error.message?.includes('429') && retryCount < 3) {
        console.log(`Rate limited, retrying in ${Math.pow(2, retryCount)} seconds...`);
        await new Promise(resolve => setTimeout(resolve, Math.pow(2, retryCount) * 1000));
        return makeRequest(retryCount + 1);
      }
      throw error;
    }
  };

  try {
    const data = await makeRequest();
    
    return {
      messages: data._results,
      nextPage: data._pagination?.next ?? null,
      hasMoreRecent: true
    };
  } catch (error: any) {
    console.error(`Error fetching messages for conversation ${conversationId}:`, error.message);
    throw error;
  }
}

function formatRecipient(recipient: { handle: string; role: string }): string {
  return recipient.handle || "Unknown";
}

function formatAuthor(author: { id: string; email: string; username: string; first_name: string } | undefined): string {
  if (!author) return "Unknown";
  return author.first_name || author.email || author.username || "Unknown";
}

function parseEmailContent(content: string): string {
  if (!content) {
    return "";
  }

  try {
    const result = talon.quotations.extractFromPlain(content);
    
    return result.body.trim();
  } catch (error) {
    console.warn("Failed to parse email content with TalonJS, using original content:", error);
    return content;
  }
}

function formatContent(message: Message): string {
  if (message.text) {
    return parseEmailContent(message.text);
  } else if (message.body) {
    return parseEmailContent(message.body);
  }
  
  return "";
}

function createTimelineFromMessages(messages: Message[]): TimelineEntry[] {
  const timeline = messages.map((message) => ({
    item: message,
    type: "message" as const,
    timestamp: message.created_at,
  })).sort((a, b) => a.timestamp - b.timestamp);

  return timeline.map(({ item, type, timestamp }) => {
    const message = item as Message;
    
    const allRecipients = message.recipients.map((recipient) => 
      formatRecipient(recipient)
    );

    const content = formatContent(message);

    return {
      timestamp: safeTimestampToDate(message.created_at),
      type: message.is_inbound ? "RECEIVED" : "SENT",
      sender: formatAuthor(message.author),
      recipients: allRecipients,
      subject: message.blurb,
      content: content,
      attachments: message.attachments?.map((att) => att.filename) || [],
    };
  });
}

function timelineToLLMFormat(timeline: TimelineEntry[]): string {
  if (timeline.length === 0) {
    return "<conversation>\nTOTAL_MESSAGES: 0\nNo messages found\n</conversation>";
  }

  const metadata = `<conversation>
TOTAL_MESSAGES: ${timeline.length}
CONVERSATION_START: ${timeline[0].timestamp.toISOString()}
CONVERSATION_END: ${timeline[timeline.length - 1].timestamp.toISOString()}
PARTICIPANTS: ${Array.from(new Set(timeline.flatMap((e) => [e.sender, ...e.recipients]))).join(", ")}
</conversation>\n\n`;

  const entries = timeline
    .map((entry, index) => {
      const timestamp = entry.timestamp.toISOString();
      const attachmentInfo =
        entry.attachments.length > 0
          ? `ATTACHMENTS:\n${entry.attachments.map((a) => `- ${a}`).join("\n")}`
          : "";

      return `<entry index="${index + 1}" type="${entry.type}">
FROM: ${entry.sender}
TO: ${entry.recipients.join(", ")}
TIMESTAMP: ${timestamp}
${entry.subject ? `SUBJECT: ${entry.subject}\n` : ""}CONTENT:
${entry.content}
${attachmentInfo}
</entry>`;
    })
    .join("\n\n");

  return metadata + entries;
}

async function main() {
  if (FRONT_FILTER) {
    console.log(`Inbox filter: ${FRONT_FILTER}`);
  } else {
    console.log(`Processing all inboxes`);
  }
  console.log(`Target datasource: ${DUST_DATASOURCE_ID}`);
  
  const cutoffDate = safeTimestampToISO(cutoffTimestamp);
  console.log(`Time filtering: Processing conversations updated after ${cutoffDate} (${DAYS_BACK} days back)`);
  console.log(`Message count filtering: Minimum ${MIN_MESSAGE_COUNT} messages per conversation`);

  try {
    const dustLimiter = new Bottleneck({
      maxConcurrent: 1,
      minTime: 60000 / DUST_TPM,
    });

    const frontLimiter = new Bottleneck({
      maxConcurrent: FRONT_MAX_CONCURRENT,
      minTime: 60000 / FRONT_TPM,
    });

    let targetInboxId: string | null = null;
    if (FRONT_FILTER) {
      try {
        const searchParams = new URLSearchParams(FRONT_FILTER);
        if (searchParams.has('inbox')) {
          const rawInboxId = searchParams.get('inbox');
          if (rawInboxId) {
            try {
              targetInboxId = convertUrlInboxIdToApiId(rawInboxId);
              console.log(`Converted inbox ID: ${rawInboxId} → ${targetInboxId}`);
            } catch (conversionError: any) {
              console.error(`Error converting inbox ID "${rawInboxId}":`, conversionError.message);
              process.exit(1);
            }
          }
        }
      } catch (parseError) {
        console.warn(`Warning: Could not parse inbox filter from "${FRONT_FILTER}"`);
      }
    }

    console.log("Fetching inboxes from Front...");
    const allInboxes = await frontLimiter.schedule(() => getAllInboxes());
    
    let inboxesToProcess: Inbox[] = [];
    if (targetInboxId) {
      const targetInbox = allInboxes.find(inbox => inbox.id === targetInboxId);
      if (targetInbox) {
        inboxesToProcess = [targetInbox];
        console.log(`Processing specific inbox: ${targetInbox.name} (${targetInbox.id})`);
      } else {
        console.warn(`Inbox with ID "${targetInboxId}" not found. Available inboxes:`);
        allInboxes.forEach(inbox => console.log(`  - ${inbox.name} (${inbox.id})`));
        return;
      }
    } else {
      inboxesToProcess = allInboxes;
      console.log(`Processing all ${allInboxes.length} inboxes`);
    }

    console.log("Starting message processing...");
    
    let totalProcessed = 0;
    let totalSuccess = 0;
    let totalErrors = 0;
    let totalSkipped = 0;

    for (const inbox of inboxesToProcess) {
      console.log(`\nProcessing inbox: ${inbox.name} (${inbox.id})`);
      
      let hasMoreConversations = true;
      let nextConversationPage: string | null = null;
      let conversationCount = 0;

      while (hasMoreConversations) {
        try {
          const { conversations, nextPage: batchNextPage, hasMoreRecent } = await frontLimiter.schedule(() => 
            getConversationsForInbox(inbox.id, nextConversationPage, BATCH_SIZE)
          );

          if (!hasMoreRecent && conversations.length === 0) {
            console.log(`Reached time cutoff (${DAYS_BACK} days back), stopping conversation processing for inbox ${inbox.name}`);
            break;
          }

          console.log(`Processing batch of ${conversations.length} conversations...`);

          for (const conversation of conversations) {
            conversationCount++;
            console.log(`Processing conversation ${conversationCount}: ${conversation.subject || 'No Subject'} (${conversation.id})`);
            
            let hasMoreMessages = true;
            let nextMessagePage: string | null = null;
            let allMessages: Message[] = [];

            while (hasMoreMessages) {
              try {
                const { messages, nextPage: messageNextPage, hasMoreRecent: messageHasMoreRecent } = await frontLimiter.schedule(() => 
                  getMessagesForConversation(conversation.id, nextMessagePage, BATCH_SIZE)
                );

                allMessages = allMessages.concat(messages);

                if (messages.length < BATCH_SIZE) {
                  hasMoreMessages = false;
                } else {
                  nextMessagePage = messageNextPage;
                  if (!nextMessagePage) {
                    hasMoreMessages = false;
                  }
                }

              } catch (error) {
                console.error(`Error collecting messages for conversation ${conversation.id}:`, error);
                console.error("Stopping script due to error");
                process.exit(1);
              }
            }

            // Check if conversation has enough messages
            if (allMessages.length < MIN_MESSAGE_COUNT) {
              console.log(`Skipping conversation ${conversation.id}: ${allMessages.length} messages (minimum ${MIN_MESSAGE_COUNT} required)`);
              totalSkipped += 1;
              continue;
            }

            try {
              await dustLimiter.schedule(() => upsertConversationToDust(conversation, allMessages));
              totalProcessed += allMessages.length;
              totalSuccess += 1;
                            
              console.log(`Completed conversation ${conversation.id}: ${allMessages.length} messages uploaded`);
            } catch (error) {
              console.error(`Error uploading conversation ${conversation.id}:`, error);
              totalErrors += 1;
              throw error;
            }
          }

          if (conversations.length < BATCH_SIZE) {
            hasMoreConversations = false;
          } else {
            nextConversationPage = batchNextPage;
            if (!nextConversationPage) {
              hasMoreConversations = false;
            }
          }

          await new Promise(resolve => setTimeout(resolve, 1000));

        } catch (error) {
          console.error(`Error processing conversations for inbox ${inbox.id}:`, error);
          console.error("Stopping script due to error");
          process.exit(1);
        }
      }

      console.log(`Completed inbox ${inbox.name}: ${conversationCount} conversations processed`);
    }

    console.log("\nImport completed!");
    console.log(`Final Summary: ${totalProcessed} messages processed`);
    console.log(`Success: ${totalSuccess} conversations imported`);
    console.log(`Skipped: ${totalSkipped} conversations (insufficient messages)`);
    console.log(`Errors: ${totalErrors} conversations failed`);
    
    if (currentRateLimit) {
      console.log(`Final Rate Limit Status: ${currentRateLimit.remaining}/${currentRateLimit.limit} requests remaining`);
    }
    
    if (totalErrors > 0) {
      console.log("Some conversations failed to import. Check the logs above for details.");
    }

  } catch (error: any) {
    console.error("Import failed:", error.message);
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