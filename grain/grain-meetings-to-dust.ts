import axios from "axios";
import * as dotenv from "dotenv";
import Bottleneck from "bottleneck";

dotenv.config();

// --- Constants --- //
const GRAIN_RATE_LIMIT_PER_MINUTE = 30;

// Required env vars
const requiredEnvVars = [
  "GRAIN_API_TOKEN",
  "DUST_API_KEY",
  "DUST_WORKSPACE_ID",
  "DUST_DATASOURCE_ID",
  "DUST_RATE_LIMIT",
  "GRAIN_MAX_CONCURRENT",
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

// --- Env config --- //
const GRAIN_START_DATE = process.env.GRAIN_START_DATE;
const GRAIN_END_DATE = process.env.GRAIN_END_DATE;
const GRAIN_API_TOKEN = process.env.GRAIN_API_TOKEN!;
const DUST_API_KEY = process.env.DUST_API_KEY!;
const DUST_WORKSPACE_ID = process.env.DUST_WORKSPACE_ID!;
const DUST_DATASOURCE_ID = process.env.DUST_DATASOURCE_ID!;

const DUST_RATE_LIMIT = parseInt(process.env.DUST_RATE_LIMIT as string);
const GRAIN_MAX_CONCURRENT = parseInt(
  process.env.GRAIN_MAX_CONCURRENT as string
);

// --- API Clients --- //
const grainApi = axios.create({
  baseURL: "https://api.grain.com",
  headers: {
    Authorization: `Bearer ${GRAIN_API_TOKEN}`,
    "Content-Type": "application/json",
  },
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

// --- Rate Limiters --- //
const grainLimiter = new Bottleneck({
  maxConcurrent: GRAIN_MAX_CONCURRENT,
  minTime: (60 * 1000) / GRAIN_RATE_LIMIT_PER_MINUTE,
});

// --- Types --- //
interface Recording {
  id: string;
  title: string;
  url: string;
  start_datetime: string;
  end_datetime: string;
  public_thumbnail_url: string | null;
}

interface Participant {
  email: string;
  name: string;
  scope: string;
}

interface Highlight {
  id: string;
  text: string;
  transcript: string;
  speakers: string[];
  timestamp: number;
  duration: number;
  created_datetime: string;
  url: string;
  thumbnail_url?: string;
  tags?: string[];
}

interface RecordingDetail extends Recording {
  participants: Participant[];
  owners: string[];
  tags: string[];
  highlights: Highlight[];
  transcript_json?: any;
  intelligence_notes_md?: any;
}

interface Section {
  prefix?: string | null;
  content?: string | null;
  sections: Section[];
}

function isWithinDateRange(
  dateStr: string,
  start?: string,
  end?: string
): boolean {
  const date = new Date(dateStr);
  if (start && date < new Date(start)) return false;
  if (end && date > new Date(end)) return false;
  return true;
}

async function getAllRecordings(): Promise<Recording[]> {
  let cursor: string | undefined = undefined;
  const allRecordings: Recording[] = [];
  let page = 1;

  do {
    const params: Record<string, any> = {
      include_participants: true,
      include_highlights: true,
      cursor,
    };
    console.log(`[Grain] Fetching page ${page}...`);
    const response = await grainLimiter.schedule(() =>
      grainApi.get("/_/public-api/recordings", { params })
    );
    const { recordings, cursor: nextCursor } = response.data;
    console.log(
      `[Grain] Page ${page}: fetched ${recordings?.length ?? 0} recordings`
    );
    allRecordings.push(...(recordings || []));
    cursor = nextCursor;
    page += 1;
  } while (cursor);

  // Date filtering
  const filteredRecordings = allRecordings.filter((rec) =>
    isWithinDateRange(rec.start_datetime, GRAIN_START_DATE, GRAIN_END_DATE)
  );

  console.log(
    `Fetched ${filteredRecordings.length} recordings from Grain (after date filtering)`
  );
  return filteredRecordings;
}

// --- Fetch Full Recording Detail --- //
async function getRecordingDetail(id: string): Promise<RecordingDetail> {
  const params = {
    include_participants: true,
    include_highlights: true,
    include_owners: true,
    transcript_format: "json", // will include transcript_json in the response
    intelligence_notes_format: "md", // will include intelligence_notes_json
  };
  const response = await grainLimiter.schedule(() =>
    grainApi.get(`/_/public-api/recordings/${id}`, { params })
  );
  return response.data;
}

// --- Format Section --- //
function formatRecordingSection(recording: RecordingDetail): Section {
  const transcriptSection = recording.transcript_json
    ? [
        {
          prefix: "Transcript",
          content: JSON.stringify(recording.transcript_json, null, 2),
          sections: [],
        },
      ]
    : [];

  const intelligenceNotesSection = recording.intelligence_notes_md
    ? [
        {
          prefix: "Intelligence Notes",
          content: recording.intelligence_notes_md,
          sections: [],
        },
      ]
    : [];

  return {
    prefix: recording.id,
    content: recording.title,
    sections: [
      {
        prefix: "Metadata",
        content: null,
        sections: [
          // ... (keep your existing metadata sections here)
        ],
      },
      // ... (existing highlights section)
      ...transcriptSection,
      ...intelligenceNotesSection,
    ],
  };
}

// --- Upsert to Dust --- //
async function upsertToDustDatasource(recording: RecordingDetail) {
  try {
    const documentId = `grain-recording-${recording.id}`;
    const section = formatRecordingSection(recording);

    await dustApi.post(
      `/w/${DUST_WORKSPACE_ID}/data_sources/${DUST_DATASOURCE_ID}/documents/${documentId}`,
      {
        section,
        source_url: recording.url,
        title: `Grain Recording ${recording.id}: ${recording.title}`,
      }
    );

    console.log(`Upserted recording ${recording.id} to Dust`);
  } catch (error: any) {
    if (axios.isAxiosError(error)) {
      console.error(
        `Error upserting recording ${recording.id} to Dust: ${error.message}\n` +
          `Status: ${error.response?.status}\n` +
          `Response data: ${JSON.stringify(error.response?.data)}`
      );
    } else {
      console.error(
        `Error upserting recording ${recording.id} to Dust:`,
        error
      );
    }
    throw error;
  }
}

// --- Main --- //
async function main() {
  try {
    const recordings = await getAllRecordings();
    // fetch details in parallel with concurrency/rate limit
    const limiter = new Bottleneck({
      maxConcurrent: GRAIN_MAX_CONCURRENT,
      minTime: (60 * 1000) / DUST_RATE_LIMIT,
    });

    const details = await Promise.all(
      recordings.map((rec) =>
        limiter.schedule(() => getRecordingDetail(rec.id))
      )
    );

    // Upsert all to Dust
    const upsertLimiter = new Bottleneck({
      maxConcurrent: GRAIN_MAX_CONCURRENT,
      minTime: (60 * 1000) / DUST_RATE_LIMIT,
    });

    const tasks = details.map((detail) =>
      upsertLimiter.schedule(() => upsertToDustDatasource(detail))
    );

    await Promise.all(tasks);

    console.log("All Grain recordings processed successfully.");
  } catch (error) {
    console.error("An error occurred:", error);
  }
}

main();
