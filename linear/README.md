# Linear to Dust Data Sync

This script is a Node.js application designed to sync Linear issue data into Dust datasources. It fetches issues from Linear based on configurable filters, including related data like comments, attachments, and history, then uploads this information to a specified Dust datasource.

## Usage example

Example of issue summary added to the Dust datasource:

```
Issue Summary for ENG-123: Implement new authentication flow

Metadata:
Issue Details:
ID: abc123
Number: 123
Identifier: ENG-123
URL: https://linear.app/company/issue/ENG-123

Team & Project:
Team: Engineering (ENG)
Project: Authentication Overhaul
State: In Progress (Active)

Dates & Times:
Created: 2024-03-18T10:00:00.000Z
Updated: 2024-03-18T15:30:00.000Z
Started: 2024-03-18T11:00:00.000Z
Due Date: 2024-03-25

People:
Creator: Jane Smith (jane@company.com)
Assignee: John Doe (john@company.com)
Subscribers: Alice Brown (alice@company.com), Bob Wilson (bob@company.com)

Planning:
Priority: 2 (High)
Estimate: 5 points
Cycle: Sprint 45 (2024-03-18T00:00:00.000Z to 2024-03-29T23:59:59.999Z)
Labels: security, authentication

Description:
Implement new OAuth2-based authentication flow with support for multiple providers.
- Add OAuth2 client implementation
- Support Google and GitHub providers
- Implement token refresh logic
- Add user session management

Comments:
Comment by Alice Brown (alice@company.com) - 2024-03-18T12:00:00.000Z
Should we also consider adding Microsoft OAuth support?
Reactions: 👍: 2, 💡: 1

Comment by John Doe (john@company.com) - 2024-03-18T13:15:00.000Z
Good suggestion. I'll create a follow-up ticket for that.

Recent History:
2024-03-18T11:00:00.000Z - John Doe
Changed status from "Todo" to "In Progress"
```

## Features

- Configurable issue filtering by:
  - Update date
  - Team
  - Project
  - State
  - Labels
- Comprehensive data collection including:
  - Comments and reactions
  - Attachments
  - Labels
  - Issue relations
  - Issue history
  - Subscribers
  - Parent/child hierarchy
  - Cycle information
  - Organization details
- Rate limiting for both Linear and Dust APIs
- Detailed logging and error handling

## Prerequisites

- Node.js (version 14 or higher)
- npm (Node Package Manager)
- Linear account with API access
- Dust account with API access
- Create a folder and fetch its data source id (starts with dts_)

## Installation

1. Clone the repository:
   ```bash
   git clone git@github.com:dust-tt/dust-labs.git
   cd dust-labs
   cd linear
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Create a `.env` file with the following variables:

   ```
   # Required variables
   LINEAR_API_KEY=your_linear_api_key
   DUST_API_KEY=your_dust_api_key
   DUST_WORKSPACE_ID=your_workspace_id
   DUST_DATASOURCE_ID=your_datasource_id
   DUST_RATE_LIMIT=100
   LINEAR_MAX_CONCURRENT=5

   # Optional filters
   LINEAR_UPDATED_SINCE=2024-01-01
   LINEAR_TEAM_KEY=TEAM
   LINEAR_PROJECT_ID=project_id
   LINEAR_STATE=Done
   LINEAR_LABEL=Bug

   # Optional data fetching configuration
   FETCH_COMMENTS=true
   FETCH_ATTACHMENTS=true
   FETCH_LABELS=true
   FETCH_RELATIONS=true
   FETCH_HISTORY=true
   FETCH_SUBSCRIBERS=true
   FETCH_HIERARCHY=true
   FETCH_CYCLE=true
   FETCH_ORGANIZATION=true
   ```

## Usage

To run the script:

```bash
npm run issues
```

The script will:
1. Validate environment variables and configuration
2. Fetch Linear issues based on configured filters
3. Process each issue and its related data
4. Upload formatted data to your Dust datasource

## How It Works

1. The script connects to Linear using the provided API key
2. It fetches issues based on the configured filters (update date, team, project, state, labels)
3. For each issue, it collects related data based on the enabled fetch configuration
4. The data is formatted into a hierarchical document structure
5. The formatted documents are upserted to the specified Dust datasource
6. Rate limiting is applied to respect both Linear and Dust API constraints

## Error Handling

The script includes comprehensive error handling:
- Validates required environment variables
- Handles API rate limits
- Provides detailed error logging
- Continues processing on individual issue failures

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the ISC License.
