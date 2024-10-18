# Dust to Notion synchronisation

This script synchronises the list of Dust assistants in a Notion database.

## Installation

1. Ensure you have Node.js (version 14 or higher) and npm installed on your system.

2. Clone this repository:
   ```
   git clone git@github.com:dust-tt/dust-labs.git
   cd dust-labs/notion
   ```

3. Install the dependencies:
   ```
   npm install --include=dev
   ```

## Environment Setup

Create a `.env` file in the root directory of the project with the following variables:

```
# source
DUST_API_KEY=your_dust_api_key
DUST_WORKSPACE_ID=your_dust_workspace_id

# destination
NOTION_API_KEY=your_notion_api_key
NOTION_DATABASE_ID=your_notion_database_id
```

Replace the placeholder values with your actual Dust and Notion settings.

## Usage

To run the script:

*To synchronise the list of Dust assistants in a Notion database:*
```
npm run sync-list-of-assistants-to-notion-database
```

## How It Works

1. The script fetches all the assistants in a given Dust workspace.
2. It configures the Notion database with some predefined properties (see [Notes](#notes) bellow)
3. It upserts each assistant to the Notion database, using the `dust.name` property as the unique identifier.

## Dependencies

- axios: For making HTTP requests to Zendesk and Dust APIs
- dotenv: For loading environment variables

## Dev Dependencies

- @types/node: TypeScript definitions for Node.js
- tsx: For running TypeScript files directly
- typescript: The TypeScript compiler

## Notes

The Notion database must be created manually.
We expect the database to be dedicated to this script. Therefore, the script will automatically configure the fields of the database.

## License

This project is licensed under the ISC License.
