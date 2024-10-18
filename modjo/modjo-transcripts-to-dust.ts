import axios, { AxiosResponse } from "axios";
import * as dotenv from "dotenv";
import Bottleneck from "bottleneck";

dotenv.config();

const MODJO_BASE_URL = process.env.MODJO_BASE_URL || "https://api.modjo.ai";
const MODJO_API_KEY = process.env.MODJO_API_KEY;
const DUST_API_KEY = process.env.DUST_API_KEY;
const DUST_WORKSPACE_ID = process.env.DUST_WORKSPACE_ID;
const DUST_VAULT_ID = process.env.DUST_VAULT_ID;
const DUST_DATASOURCE_ID = process.env.DUST_DATASOURCE_ID;
const INCLUDE_CONTACT_DETAILS = process.env.INCLUDE_CONTACT_DETAILS !== 'false';
const INCLUDE_RECORDING_URL = process.env.INCLUDE_RECORDING_URL !== 'false';

if (
  !MODJO_API_KEY ||
  !DUST_API_KEY ||
  !DUST_WORKSPACE_ID ||
  !DUST_VAULT_ID ||
  !DUST_DATASOURCE_ID
) {
  throw new Error(
    "Please provide values for MODJO_API_KEY, DUST_API_KEY, DUST_WORKSPACE_ID, DUST_VAULT_ID, and DUST_DATASOURCE_ID in .env file."
  );
}

// Can be `null` if you want to fetch all transcripts
const YESTERDAY = new Date(Date.now() - 86400000).toISOString().split('T')[0];
const TRANSCRIPTS_SINCE = process.env.TRANSCRIPTS_SINCE === "null" ? null : (process.env.TRANSCRIPTS_SINCE || YESTERDAY);

const modjoApi = axios.create({
  baseURL: MODJO_BASE_URL,
  headers: {
    "Content-Type": "application/json",
    "X-API-KEY": MODJO_API_KEY,
  },
  maxContentLength: Infinity,
  maxBodyLength: Infinity,
});

const dustApi = axios.create({
  baseURL: "https://dust.tt/api/v1",
  headers: {
    Authorization: `Bearer ${DUST_API_KEY}`,
    "Content-Type": "application/json",
  },
  maxContentLength: Infinity,
  maxBodyLength: Infinity,
});

// Set up rate limiting for Dust API calls
const limiter = new Bottleneck({
  minTime: 500, // 500ms between tasks
  maxConcurrent: 1, // Run 1 task at a time
});

interface ModjoCallExport {
  callId: number;
  title: string;
  startDate: string;
  duration: number;
  provider: string;
  language: string;
  callCrmId: string | null;
  relations: {
    recording: {
      url: string;
    };
    highlights: {
      content: string;
    } | null;
    speakers: {
      contactId?: number;
      userId?: number;
      email: string | null;
      name: string;
      phoneNumber?: string | null;
      speakerId: number;
      type: string;
    }[];
    transcript: {
      startTime: number;
      endTime: number;
      speakerId: number;
      content: string;
      topics: { topicId: number; name: string }[];
    }[];
    tags: {
      name: string;
    }[];
  };
}

async function getModjoTranscripts(): Promise<ModjoCallExport[]> {
  let allTranscripts: ModjoCallExport[] = [];
  let page = 1;
  const perPage = 50;

  console.log(
    `Will retrieve all transcripts since: ${TRANSCRIPTS_SINCE}\n` +
    `Will include contact details: ${INCLUDE_CONTACT_DETAILS}\n` +
    `Will include recording URLs: ${INCLUDE_RECORDING_URL}`
  );

  do {
    try {
      const response: AxiosResponse<{
        pagination: { totalValues: number; lastPage: number };
        values: ModjoCallExport[];
      }> = await modjoApi.post("/v1/calls/exports", {
        pagination: { page, perPage },
        filters: {
          callStartDateRange: TRANSCRIPTS_SINCE
            ? {
                start: `${TRANSCRIPTS_SINCE}T00:00:00Z`,
                end: new Date().toISOString(),
              }
            : undefined,
        },
        relations: {
          recording: true,
          highlights: true,
          transcript: true,
          speakers: true,
          tags: true,
        },
      });

      allTranscripts = allTranscripts.concat(response.data.values);
      console.log(
        `Retrieved ${response.data.values.length} transcripts. Total: ${allTranscripts.length}`
      );

      if (page >= response.data.pagination.lastPage) {
        break;
      }
      page++;
    } catch (error) {
      if (axios.isAxiosError(error) && error.response) {
        console.error("Error fetching Modjo transcripts:", error.response.data);
      } else {
        console.error("Error fetching Modjo transcripts:", error);
      }
      break;
    }
  } while (true);

  return allTranscripts;
}

async function upsertToDustDatasource(transcript: ModjoCallExport) {
  const documentId = `modjo-transcript-${transcript.callId}`;

  let content = `Call ID: ${transcript.callId}\n`;
  content += `Tags: ${transcript.relations.tags.map(tag => `"${tag.name}"`).join(', ') }\n`;
  content += `Title: ${transcript.title}\n`;
  content += `Date: ${transcript.startDate}\n`;
  content += `Duration: ${transcript.duration} seconds\n`;
  content += `Provider: ${transcript.provider}\n`;
  content += `Language: ${transcript.language}\n`;
  if (transcript.callCrmId) content += `CRM ID: ${transcript.callCrmId}\n`;
  if (INCLUDE_RECORDING_URL && transcript.relations.recording)
    content += `Recording URL: ${transcript.relations.recording.url}\n`;


  content += "\n# Speakers\n";
  transcript.relations.speakers.forEach((speaker) => {
    content += `${speaker.speakerId}: ${speaker.name} (${speaker.type})`;
    if (INCLUDE_CONTACT_DETAILS) {
      if (speaker.email) content += ` - Email: ${speaker.email}`;
      if (speaker.phoneNumber) content += ` - Phone: ${speaker.phoneNumber}`;
    }
    content += "\n";
  });

  if (transcript.relations.highlights)
    content += `\n# Highlights\n${transcript.relations.highlights.content.trim()}\n`;

  content += "\n# Transcript\n";
  transcript.relations.transcript.forEach((entry) => {
    const speaker = transcript.relations.speakers.find(
      (s) => s.speakerId === entry.speakerId
    );
    const speakerName = speaker ? speaker.name : `Speaker ${entry.speakerId}`;
    content += `[${formatTime(entry.startTime)} - ${formatTime(
      entry.endTime
    )}] ${speakerName}: ${entry.content}\n`;
    if (entry.topics.length > 0) {
      content += `Topics: ${entry.topics.map((t) => t.name).join(", ")}\n`;
    }
    content += "\n";
  });

  try {
    await limiter.schedule(() =>
      dustApi.post(
        `/w/${DUST_WORKSPACE_ID}/vaults/${DUST_VAULT_ID}/data_sources/${DUST_DATASOURCE_ID}/documents/${documentId}`,
        {
          text: content.trim(),
          source_url: `https://app.modjo.ai/call-details/${transcript.callId}`,
        }
      )
    );
    console.log(`Upserted transcript ${transcript.callId} to Dust datasource`);
  } catch (error) {
    if (axios.isAxiosError(error) && error.response) {
      console.error(
        `Error upserting transcript ${transcript.callId} to Dust datasource:`,
        error.response.data
      );
    } else {
      console.error(
        `Error upserting transcript ${transcript.callId} to Dust datasource:`,
        error
      );
    }
  }
}

function formatTime(seconds: number): string {
  const minutes = Math.floor(seconds / 60);
  const remainingSeconds = Math.floor(seconds % 60);
  return `${minutes.toString().padStart(2, "0")}:${remainingSeconds
    .toString()
    .padStart(2, "0")}`;
}

async function main() {
  try {
    const transcripts = await getModjoTranscripts();
    console.log(`Found ${transcripts.length} transcripts.`);

    for (const transcript of transcripts) {
      await upsertToDustDatasource(transcript);
    }

    console.log("All transcripts processed successfully.");
  } catch (error) {
    console.error("An error occurred:", error);
  }
}

main();
