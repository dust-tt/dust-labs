# DEPRECATION WARNING

This script is **DEPRECATED** and will not be maintained anymore.
Please use the [Salesforce native tool](https://docs.dust.tt/docs/salesforce) if you need to interact with Salesforce in Dust.

# Salesforce to Dust Data Sync

This script is a multi-threaded Node.js application designed to import Salesforce account data into Dust datasources. It fetches recently updated accounts from Salesforce, including related contacts, opportunities, and cases, and then uploads this information to a specified Dust datasource.

## Usage example

![Example usage of the script](https://i.ibb.co/rtfTyjH/Screenshot-2024-08-20-at-12-00-24.png)

Example of account summary added to the Dust datasource:

```
Account Summary for Grand Hotels & Resorts Ltd dd
Basic Account Details:
Company Name: Grand Hotels & Resorts Ltd dd
Industry: Hospitality
Annual Revenue: $500,000,000
Number of Employees: 5600
Phone: (312) 596-1000
Website: www.grandhotels.com
Locations:
Billing Address: 2334 N. Michigan Avenue, Suite 1500
Chicago, IL 60601, USA, Chicago, IL
Shipping Address: 2334 N. Michigan Avenue, Suite 1500
Chicago, IL 60601, USA
Key Contacts:
- John Bond, VP, Facilities
    Email: bond_john@grandhotels.com, Phone: (312) 596-1000
    Last Modified: 2022-10-17T08:54:06.000+0000
- Tim Barr, SVP, Administration and Finance
    Email: barr_tim@grandhotels.com, Phone: (312) 596-1000
    Last Modified: 2022-10-17T08:54:06.000+0000
Account Status:
Type: Customer - Direct
Rating: Warm
Account Source: N/A
Created Date: 2022-10-17T08:54:06.000+0000
Last Modified Date: 2023-11-30T08:45:06.000+0000
Last Activity Date: N/A
Sales Information:
Open Opportunities:
- Grand Hotels SLA
    Stage: Closed Won, Amount: $90,000, Close Date: 2022-07-06
    Last Modified: 2023-12-18T15:28:14.000+0000
- Grand Hotels Generator Installations
    Stage: Closed Won, Amount: $350,000, Close Date: 2022-09-25
    Last Modified: 2022-10-17T08:54:06.000+0000
- GRAND Hotels Guest Portable Generators TEAMS UPDATED
    Stage: Proposal/Price Quote, Amount: $2,500, Close Date: 2023-03-16
    Last Modified: 2024-02-29T22:31:41.000+0000
- Grand Hotels Emergency Generators
    Stage: Closed Won, Amount: $210,000, Close Date: 2022-09-23
    Last Modified: 2022-10-17T08:54:06.000+0000
- Grand Hotels Kitchen Generator
    Stage: Value Proposition, Amount: $1,200, Close Date: 2023-01-18
    Last Modified: 2024-08-20T07:14:36.000+0000
Account Health:
Recent Support Cases:
- Case Number: 00001013
    Subject: Starting up generator consumes excessive power, Status: Closed, Created Date: 2022-10-17T08:54:06.000+0000
    Last Modified: 2022-10-17T08:54:06.000+0000
- Case Number: 00001008
    Subject: Customer service for portable generators needs beefing up, Status: Closed, Created Date: 2022-10-17T08:54:06.000+0000
    Last Modified: 2022-10-17T08:54:06.000+0000
- Case Number: 00001014
    Subject: Delay in installation; spare parts unavailable, Status: Closed, Created Date: 2022-10-17T08:54:06.000+0000
    Last Modified: 2022-10-17T08:54:06.000+0000
- Case Number: 00001007
    Subject: Structural breakdown of rotor assembly, Status: Closed, Created Date: 2022-10-17T08:54:06.000+0000
    Last Modified: 2022-10-17T08:54:06.000+0000
Additional Information:
Chain of hotels and resorts across the US, UK, Eastern Europe, Japan, and SE Asia.
```

## Features

- Multi-threaded processing for improved performance
- Fetches accounts updated within the last 24 hours
- Includes related Salesforce objects (Contacts, Opportunities, Cases)
- Supports both username-password and OAuth 2.0 authentication with Salesforce
- Upserts data to Dust datasources
- Passing the env variable `IMPORT_AS_TABLE` will also save accounts in a structured tables in Dust.

## Prerequisites

- Node.js (version 14 or higher recommended)
- npm (Node Package Manager)
- Salesforce account with API access
- Dust account with API access

## Installation

1. Clone the repository:

   ```
   git clone https://github.com/your-username/salesforce-to-dust-import.git
   cd salesforce-to-dust-import
   ```

2. Install dependencies:

   ```
   npm install
   ```

3. Create a `.env` file in the root directory with the following variables:

   ```
   SF_LOGIN_URL=https://login.salesforce.com
   SF_USERNAME=your_salesforce_username
   SF_PASSWORD=your_salesforce_password
   SF_SECURITY_TOKEN=your_salesforce_security_token
   SF_CLIENT_ID=your_salesforce_client_id
   SF_CLIENT_SECRET=your_salesforce_client_secret
   DUST_API_KEY=your_dust_api_key
   DUST_WORKSPACE_ID=your_dust_workspace_id
   DUST_VAULT_ID=your_dust_vault_id
   DUST_DATASOURCE_ID_SALESFORCE=your_dust_datasource_id
   ```

   Note: You can use either username-password authentication (SF_USERNAME, SF_PASSWORD, SF_SECURITY_TOKEN) or OAuth 2.0 (SF_CLIENT_ID, SF_CLIENT_SECRET) for Salesforce.

## Usage

To run the script:

```
npm run accounts
```

This command will execute the `sfdc-accounts-to-dust.ts` script using ts-node.

## Configuration

You can adjust the following variables in the script:

- `UPDATED_SINCE`: Change the time range for fetching updated accounts (default is last 24 hours)
- `THREADS_NUMBER`: Adjust the number of worker threads (default is 5)

## How It Works

1. The script connects to Salesforce using the provided credentials.
2. It fetches IDs of accounts that have been updated in the last 24 hours, including accounts with updated contacts, opportunities, or cases.
3. Detailed information for these accounts is retrieved from Salesforce, including related objects.
4. The accounts are divided into batches and processed by multiple worker threads.
5. Each account's information is formatted into a structured text document.
6. The formatted documents are upserted to the specified Dust datasource.

## Error Handling

The script includes error handling for various scenarios:

- Connection errors with Salesforce or Dust API
- Query execution errors
- Data processing errors

Errors are logged to the console for debugging purposes.

## Linting

To run the linter:

```
npm run lint
```

## Building

To compile TypeScript to JavaScript:

```
npm run build
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the ISC License.
