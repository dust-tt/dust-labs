# Zendesk to Dust Datasource Import

This script imports a full Nuclino workspace into a Dust datasource. 

## Installation

1. Ensure you have Node.js (recent version, tested with Node 20) and npm installed on your system.

2. Clone this repository:
   ```
   git git@github.com:dust-tt/dust-labs.git
   cd nuclino/import-script
   ```

3. Install the dependencies:
   ```
   npm install
   ```

## Environment Setup

Create a `.env` file in the root directory of the project with the following variables:

```
NUCLINO_API_KEY=<your Nuclino API key>
DUST_API_KEY=<your Dust API key>
DUST_WORKSPACE_ID=<Dust workspace id>
```

Replace the placeholder values with your actual Nuclino and Dust credentials.

## Usage

To run the script:

*To import an entire workspace
```
npm run import <nuclinoWorkspaceName> <DustSpaceId> <DustDataSourceId> <mode>sync|archive|dryrun</mode>
```
- \<nuclinoWorkspaceName\> - The name of the Nuclino workspace to import.
- \<DustSpaceId\> - The ID of the Dust space to import into.
- \<DustDataSourceId\> - The ID of the Dust datasource to import into.
- <mode>sync|archive|dryrun</mode> - The mode to run the script in. Options are sync, archive, and dryrun.
  - archive: Will archive the documents that were not found in nuclino. Removed documents in Nuclino will remain in Dust
  - sync: Will archive the documents that were not found in nuclino. Removed documents in Nuclino will be removed from Dust as well
  - dryrun: Will not archive or remove any documents in Nuclino or Dust but output what will happen if you run it.

## How It Works

1. The script will recursively fetch all items in the root of a workspace
2. If the item is a collection, it will recurse into it. 
3. If it's an article, it will fetch the article data and
4. The collected data is formatted into a single text document.
5. The formatted data is then upserted to the specified Dust datasource, using the article Id as document Id.



## Error Handling

The script includes basic error handling, but your mileage may vary

## Building

To compile the TypeScript code to JavaScript:

```
npm run build
```

This will create a `dist` directory with the compiled JavaScript files.

## Dependencies

- axios: For making HTTP requests to Nuclino and Dust APIs
- dotenv: For loading environment variables
- Bottleneck: For limiting the number of concurrent operations

## Dev Dependencies

- @types/node: TypeScript definitions for Node.js
- ts-node: For running TypeScript files directly
- typescript: The TypeScript compiler

## Notes

- Ensure you have the necessary permissions in both Nuclino and Dust to perform these operations.
- Be mindful of your Zendesk API usage limits when running this script frequently or with large datasets.
- The script currently fetches tickets from the last 24 hours. Modify the `TICKETS_UPDATED_SINCE` constant if you need a different time range.

## License

This project is licensed under the ISC License.
