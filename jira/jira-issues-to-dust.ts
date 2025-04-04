import axios, { AxiosResponse, AxiosError } from "axios";
import * as dotenv from "dotenv";
import Bottleneck from "bottleneck";

dotenv.config();

const DEFAULT_JIRA_QUERY = "updated >= -24h ORDER BY updated DESC";

const JIRA_SUBDOMAIN = process.env.JIRA_SUBDOMAIN;
const JIRA_EMAIL = process.env.JIRA_EMAIL;
const JIRA_API_TOKEN = process.env.JIRA_API_TOKEN;
const JIRA_QUERY = process.env.JIRA_QUERY || DEFAULT_JIRA_QUERY;
const DUST_API_KEY = process.env.DUST_API_KEY;
const DUST_WORKSPACE_ID = process.env.DUST_WORKSPACE_ID;
const DUST_VAULT_ID = process.env.DUST_VAULT_ID;
const DUST_DATASOURCE_ID = process.env.DUST_DATASOURCE_ID;

const requiredEnvVars = [
  "JIRA_SUBDOMAIN",
  "JIRA_EMAIL",
  "JIRA_API_TOKEN",
  "DUST_API_KEY",
  "DUST_WORKSPACE_ID",
  "DUST_VAULT_ID",
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

const DUST_RATE_LIMIT = 120; // requests per minute

const jiraApi = axios.create({
  baseURL: `https://${JIRA_SUBDOMAIN}.atlassian.net/rest/api/3`,
  auth: {
    username: JIRA_EMAIL as string,
    password: JIRA_API_TOKEN as string,
  },
  headers: {
    "Content-Type": "application/json",
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

interface JiraIssue {
  id: string;
  key: string;
  self: string;
  fields: {
    summary: string;
    description: {
      content: Array<{
        content: Array<{
          text: string;
        }>;
      }>;
    };
    issuetype: {
      name: string;
    };
    status: {
      name: string;
    };
    priority: {
      name: string;
    };
    assignee: {
      displayName: string;
      emailAddress: string;
    } | null;
    reporter: {
      displayName: string;
      emailAddress: string;
    };
    project: {
      key: string;
      name: string;
    };
    created: string;
    updated: string;
    resolutiondate: string | null;
    resolution: {
      name: string;
    } | null;
    labels: string[];
    components: Array<{ name: string }>;
    sprint: {
      name: string;
    } | null;
    epic: {
      name: string;
    } | null;
    timeoriginalestimate: number | null;
    timeestimate: number | null;
    timespent: number | null;
    votes: {
      votes: number;
    };
    watches: {
      watchCount: number;
    };
    fixVersions: Array<{ name: string }>;
    versions: Array<{ name: string }>;
    subtasks: Array<{ key: string; fields: { summary: string } }>;
    issuelinks: Array<{
      type: { name: string; inward: string; outward: string };
      inwardIssue?: { key: string; fields: { summary: string } };
      outwardIssue?: { key: string; fields: { summary: string } };
    }>;
    attachment: Array<{
      filename: string;
      content: string;
    }>;
    comment: {
      comments: Array<{
        author: {
          displayName: string;
          emailAddress: string;
        };
        created: string;
        body: {
          content: Array<{
            content: Array<{
              text: string;
            }>;
          }>;
        };
      }>;
    };
  };
}

interface JiraSearchResponse {
  issues: JiraIssue[];
  startAt: number;
  maxResults: number;
  total: number;
}

async function getIssuesUpdatedLast24Hours(): Promise<JiraIssue[]> {
  let allIssues: JiraIssue[] = [];
  let startAt = 0;
  const maxResults = 50;
  let total = 0;

  const makeRequest = async (
    retryCount = 0
  ): Promise<AxiosResponse<JiraSearchResponse>> => {
    try {
      return await jiraApi.post("/search", {
        jql: JIRA_QUERY,
        startAt,
        maxResults,
        fields: [
          "summary",
          "description",
          "issuetype",
          "status",
          "priority",
          "assignee",
          "reporter",
          "project",
          "created",
          "updated",
          "resolutiondate",
          "resolution",
          "labels",
          "components",
          "timeoriginalestimate",
          "timeestimate",
          "timespent",
          "votes",
          "watches",
          "fixVersions",
          "versions",
          "subtasks",
          "issuelinks",
          "attachment",
          "comment",
        ],
        expand: ["renderedFields"],
      });
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
      allIssues = allIssues.concat(response.data.issues);
      total = response.data.total;
      startAt += maxResults;

      console.log(`Retrieved ${allIssues.length} of ${total} issues`);
    } catch (error) {
      if (axios.isAxiosError(error)) {
        console.error("Error fetching JIRA issues:");
        console.error("Status:", error.response?.status);
        console.error("Data:", JSON.stringify(error.response?.data, null, 2));
        console.error("Config:", JSON.stringify(error.config, null, 2));
      } else {
        console.error("Unexpected error:", error);
      }
      break;
    }
  } while (allIssues.length < total);

  console.log(`Final total: ${allIssues.length} issues retrieved`);
  return allIssues;
}

function formatDescription(
  description: JiraIssue["fields"]["description"]
): string {
  return (
    description?.content
      .map((c) => c.content?.map((cc) => cc.text).join(""))
      .join("\n") || ""
  );
}

function formatComments(
  comments: JiraIssue["fields"]["comment"]["comments"]
): string {
  return comments
    .map(
      (comment) => `
[${comment.created}] Author: ${comment.author.displayName} (${
        comment.author.emailAddress
      })
${comment.body.content
  .map((c) => c.content.map((cc) => cc.text).join(""))
  .join("\n")}
`
    )
    .join("\n");
}

async function upsertToDustDatasource(issue: JiraIssue) {
  const documentId = `issue-${issue.key}`;
  const content = `
Issue Key: ${issue.key}
ID: ${issue.id}
URL: ${issue.self}
Summary: ${issue.fields.summary}
Description:
${formatDescription(issue.fields.description)}

Issue Type: ${issue.fields.issuetype.name}
Status: ${issue.fields.status.name}
Priority: ${issue.fields.priority.name}
Assignee: ${
    issue.fields.assignee
      ? `${issue.fields.assignee.displayName} (${issue.fields.assignee.emailAddress})`
      : "Unassigned"
  }
Reporter: ${issue.fields.reporter.displayName} (${
    issue.fields.reporter.emailAddress
  })
Project: ${issue.fields.project.name} (${issue.fields.project.key})
Created: ${issue.fields.created}
Updated: ${issue.fields.updated}
Resolution: ${
    issue.fields.resolution ? issue.fields.resolution.name : "Unresolved"
  }
Resolution Date: ${issue.fields.resolutiondate || "N/A"}
Labels: ${issue.fields.labels.join(", ")}
Components: ${issue.fields.components.map((c) => c.name).join(", ")}
Sprint: ${issue.fields.sprint ? issue.fields.sprint.name : "N/A"}
Epic: ${issue.fields.epic ? issue.fields.epic.name : "N/A"}
Time Tracking:
  Original Estimate: ${issue.fields.timeoriginalestimate || "N/A"}
  Remaining Estimate: ${issue.fields.timeestimate || "N/A"}
  Time Spent: ${issue.fields.timespent || "N/A"}
Votes: ${issue.fields.votes.votes}
Watches: ${issue.fields.watches.watchCount}
Fix Versions: ${issue.fields.fixVersions.map((v) => v.name).join(", ")}
Affected Versions: ${issue.fields.versions.map((v) => v.name).join(", ")}
Subtasks: ${issue.fields.subtasks
    .map((st) => `${st.key}: ${st.fields.summary}`)
    .join(", ")}
Issue Links: ${issue.fields.issuelinks
    .map((link) => {
      const linkedIssue = link.inwardIssue || link.outwardIssue;
      return linkedIssue
        ? `${link.type.name} ${linkedIssue.key}: ${linkedIssue.fields.summary}`
        : "";
    })
    .filter(Boolean)
    .join(", ")}
Attachments: ${issue.fields.attachment.map((a) => a.filename).join(", ")}

Comments:
${formatComments(issue.fields.comment.comments)}
  `.trim();

  try {
    await dustApi.post(
      `/w/${DUST_WORKSPACE_ID}/vaults/${DUST_VAULT_ID}/data_sources/${DUST_DATASOURCE_ID}/documents/${documentId}`,
      {
        text: content,
      }
    );
    console.log(`Upserted issue ${issue.key} to Dust datasource`);
  } catch (error) {
    console.error(
      `Error upserting issue ${issue.key} to Dust datasource:`,
      error
    );
    throw error;
  }
}

async function main() {
  try {
    const recentIssues = await getIssuesUpdatedLast24Hours();
    console.log(
      `Found ${recentIssues.length} issues updated in the last 24 hours.`
    );

    const limiter = new Bottleneck({
      maxConcurrent: DUST_RATE_LIMIT,
      minTime: (60 * 1000) / DUST_RATE_LIMIT,
    });

    const tasks = recentIssues.map((issue) =>
      limiter.schedule(() => upsertToDustDatasource(issue))
    );
    await Promise.all(tasks);
    console.log("All issues processed successfully.");
  } catch (error) {
    console.error("An error occurred:", error);
  }
}

main();
