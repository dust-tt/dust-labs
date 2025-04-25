# Jira Issues to Dust Datasource Import Script

This script synchronizes Jira issues with a Dust datasource. It fetches recently updated Jira issues and upserts them into a specified Dust datasource, allowing for easy integration of Jira data into Dust-powered applications.

## Table of Contents

1. [Example](#example)
2. [Installation](#installation)
3. [Environment Setup](#environment-setup)
4. [Configuration](#configuration)
5. [Usage](#usage)
6. [Script Details](#script-details)
7. [Troubleshooting](#troubleshooting)

## Example
![Example usage of the script](https://i.ibb.co/t8dtM0P/Screenshot-2024-08-20-at-16-40-32.png)

Example of issues summary added to the Dust datasource: 

```
Issue Key: KAN-3
ID: 10002
URL: https://dust4ai.atlassian.net/rest/api/3/issue/10002
Summary: Nice task to do
Description:
What do you think about this description

Issue Type: Task
Status: To Do
Priority: Medium
Assignee: Alban Dumouilla (alban@dust.tt)
Reporter: Alban Dumouilla (alban@dust.tt)
Project: My Kanban Project (KAN)
Created: 2024-08-20T15:51:58.901+0200
Updated: 2024-08-20T16:26:36.215+0200
Resolution: Unresolved
Resolution Date: N/A
Labels: test
Components: 
Sprint: N/A
Epic: N/A
Time Tracking:
  Original Estimate: N/A
  Remaining Estimate: N/A
  Time Spent: N/A
Votes: 0
Watches: 1
Fix Versions: 
Affected Versions: 
Subtasks: 
Issue Links: Blocks KAN-1: This is a todo card
Attachments: 

Comments:

[2024-08-20T16:26:36.076+0200] Author: Alban Dumouilla (alban@dust.tt)
This is the first comment of the issue
```

## Installation

1. Clone this repository:
   ```bash
   git clone git@github.com:dust-tt/dust-labs.git
   cd dust-labs/jira
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

## Environment Setup

Create a `.env` file in the root directory of the project with the following variables:

```
JIRA_SUBDOMAIN=your-jira-subdomain # eg dust if your url is https://dust.atlassian.net
JIRA_EMAIL=your-jira-email
JIRA_API_TOKEN=your-jira-api-token
JIRA_QUERY=your-jira-query # optional, default is `updated >= -24h ORDER BY updated DESC`
DUST_API_KEY=your-dust-api-key
DUST_WORKSPACE_ID=your-dust-workspace-id
DUST_DATASOURCE_ID=your-dust-datasource-id # you get it from Company Data > Folders > ... > API
DUST_VAULT_ID=your-dust-space-id
```

Replace the placeholder values with your actual Jira and Dust credentials.

## Usage

Run the script using the following command:

```bash
npm run issues
```

This will execute the `jira-issues-to-dust.ts` script using `ts-node`.

## Script Details

### Functionality

1. **Fetching Jira Issues**: 
   - Retrieves all Jira issues updated within the last 24 hours (or as specified by `JIRA_QUERY`).
   - Paginates through results to ensure all issues are fetched.

2. **Processing Issues**:
   - For each issue, formats the data into a structured text representation.
   - Includes details such as issue key, summary, description, status, assignee, comments, and more.

3. **Upserting to Dust**:
   - Creates or updates a document in the specified Dust datasource for each Jira issue.
   - Uses the Jira issue key as the unique identifier for the Dust document.

4. **Concurrency**:
   - Utilizes `p-limit` to control the number of concurrent requests to Dust API.

### Key Components

- `getIssuesUpdatedLast24Hours()`: Fetches recent Jira issues.
- `formatDescription()` and `formatComments()`: Helper functions to format Jira issue data.
- `upsertToDustDatasource()`: Sends formatted issue data to Dust.
- `main()`: Orchestrates the entire process.

## Troubleshooting

- **Jira API Errors**: 
  - Ensure your Jira credentials are correct in the `.env` file.
  - Check if you have the necessary permissions to access the Jira API.

- **Dust API Errors**:
  - Verify your Dust API key, workspace ID, and datasource ID in the `.env` file.
  - Ensure the specified datasource exists and you have write permissions.

- **Rate Limiting**:
  - If you encounter rate limiting issues, try reducing the `THREADS_NUMBER` constant.

- **Memory Issues**:
  - For large Jira instances, you might need to process issues in batches. Consider modifying the script to fetch and process issues in smaller time ranges.

For any other issues, check the console output for error messages and stack traces.
