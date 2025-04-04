import * as dotenv from 'dotenv';
import Bottleneck from 'bottleneck';


/**
 * All Things Nuclino go here until I take the time to split thing up nicely.
 */

interface NuclinoField {
    object: 'field';
    id: string;
    type: string;
    name: string;
}

interface NuclinoItem {
    object: string;
    id: string;
    createdAt: Date;
    createdUserId: string;
    childIds: Array<string>;
}

class NuclinoWorkspace implements NuclinoItem {
    object: string = 'workspace';
    id: string;
    teamId: string;
    name: string;
    createdAt: Date;
    createdUserId: string;
    fields: NuclinoField[];
    childIds: string[];

    private constructor(id: string, teamId: string, name: string, createdAt: string, createdUserId: string) {
        this.id = id;
        this.teamId = teamId;
        this.name = name;
        this.createdAt = new Date(createdAt);
        this.createdUserId = createdUserId;
    }

    public static fromRawData(data: any) {
        let instance = new NuclinoWorkspace(data.id, data.teamId, data.name, data.createdAt, data.createdUserId);

        instance.fields = data.fields.map((field: any) => ({
            object: field.object,
            id: field.id,
            type: field.type,
            name: field.name,
        }));
        instance.childIds = [...data.childIds];
        return instance;
    }
}

class NuclinoUser {
    object: string = "user";
    id: string;
    firstName: string;
    lastName: string;
    email: string;

    private constructor(id: string, firstName: string, lastName: string, email: string) {
        this.id = id;
        this.firstName = firstName;
        this.lastName = lastName;
        this.email = email;
    }

    public static fromRawData(data: any) {
        return new NuclinoUser(data.id, data.firstName = 'deactivated', data.lastName = '', data.email);

    }

}

class NuclinoCollection implements NuclinoItem {
    object: string = 'collection';
    id: string;
    workspaceId: string;
    url: string;
    title: string;
    createdAt: Date;
    createdUserId: string;
    lastUpdatedAt: Date;
    lastUpdatedUserId: string;
    childIds: Array<string>;


    private constructor(id: string, workspaceId: string, title: string, createdAt: string, createdUserId: string, lastUpdatedAt: string, lastUpdatedUserId: string) {
        this.id = id;
        this.workspaceId = workspaceId;
        this.title = title;
        this.createdAt = new Date(createdAt);
        this.createdUserId = createdUserId;
        this.lastUpdatedAt = new Date(lastUpdatedAt);
        this.lastUpdatedUserId = lastUpdatedUserId;
    }

    public static fromRawData(data: any) {
        let instance = new NuclinoCollection(data.id, data.workspaceId, data.title, data.createdAt, data.createdUserId, data.lastUpdatedAt, data.lastUpdatedUserId);

        instance.childIds = [...data.childIds];
        return instance;
    }

}

class NuclinoArticle implements NuclinoItem {
    object: string = 'item';
    id: string;
    workspaceId: string;
    url: string;
    title: string;
    createdAt: Date;
    createdUserId: string;
    lastUpdatedAt: Date;
    lastUpdatedUserId: string;
    fields: Array<NuclinoField> = [];

    childIds: Array<string> = [];
    content: string;
    meta: {
        itemIds: Array<string>,
        fileIds: Array<string>
    } = {itemIds: [], fileIds: []};


    private constructor(id: string, workspaceId: string, title: string, url: string, createdAt: string, createdUserId: string, lastUpdatedAt: string, lastUpdatedUserId: string, content: string) {
        this.id = id;
        this.workspaceId = workspaceId;
        this.title = title;
        this.url = url;
        this.createdAt = new Date(createdAt);
        this.createdUserId = createdUserId;
        this.lastUpdatedAt = new Date(lastUpdatedAt);
        this.lastUpdatedUserId = lastUpdatedUserId;
        this.content = content;
    }

    public static fromRawData(data: any) {
        let instance = new NuclinoArticle(data.id, data.workspaceId, data.title, data.url, data.createdAt, data.createdUserId, data.lastUpdatedAt, data.lastUpdatedUserId, data.content);
        /*instance.fields = data.fields.map((field: any) => ({
            object: field.object,
            id: field.id,
            type: field.type,
            name: field.name,
        }));
        */

        instance.meta = data.contentMeta
        return instance;
    }
}

/*class NuclinoFile implements NuclinoItem {
    object: string;
    id: string;
    itemId: string;
    fileName: string;
    createdAt: Date;
    createdUserId: string;
    download: {
        url: string,
        expiresAt: Date
    };

    private constructor(id: string, itemId: string, fileName: string, createdAt: string, createdUserId: string, download: any) {
        this.id = id;
        this.itemId = itemId;
        this.fileName = fileName;
        this.createdAt = new Date(createdAt);
        this.createdUserId = createdUserId;
        this.download = {
            url: download.url,
            expiresAt: new Date(download.expiresAt)
        }
    }

    public static fromRawData(data: any) {
        return new NuclinoFile(data.id, data.itemId, data.fileName, data.createdAt, data.createdUserId, data.download);
    }


}
*/


class Nuclino {
    private users: Map<string, NuclinoUser> = new Map();

    constructor(
        private rateLimiter: Bottleneck,
    ) {
        this.rateLimiter = rateLimiter;
    }

    public async getUserInfo(userId: string): Promise<NuclinoUser> {
        try {
            if (this.users.has(userId)) {
                return this.users.get(userId);
            }
            const response = await this.doThrottledGetRequest(`/users/${userId}`);
            const user = NuclinoUser.fromRawData(response.data);
            this.users.set(userId, user);
            return user;
        } catch (error) {
            console.error(
                `Error fetching  user ${userId}:`,
                error.message
            );
        }

    }


    private async doThrottledGetRequest(url: string, data?: any) {
        const params = new URLSearchParams();
        if (data?.params) {
            Object.entries(data.params).forEach(([key, value]) => {
                if (typeof value === "string") {
                    params.append(key, value);
                }
            });
        }

        const queryString = params.toString();
        const fullUrl = `https://api.nuclino.com/v0${url}${queryString ? `?${queryString}` : ''}`;

        return this.rateLimiter.schedule(async () => {
            try {
                const response = await fetch(fullUrl, {
                    method: 'GET',
                    headers: {
                        'Authorization': `${NUCLINO_API_KEY}`,
                        'Accept': 'application/json'
                    },
                    signal: AbortSignal.timeout(60 * 1000), // <=== HERE
                });

                if (!response.ok) {
                    throw new Error(`HTTP error! Status: ${response.status}`);
                }

                return await response.json();
            } catch (error) {
                console.log("error", error);
                throw error;
            }
        });
    }


    public async getWorkspace(name: string): Promise<NuclinoWorkspace> {
        try {
            const response = await this.doThrottledGetRequest(`/workspaces`);
            const workspaceRawata = response.data.results.find(workspace => workspace.name === name);
            return NuclinoWorkspace.fromRawData(workspaceRawata);

        } catch (error) {
            console.error(
                `Error fetching  workspace ${name}:`,
                error.message
            );
            throw error;
        }

    }


    public async getItem(itemId: string): Promise<NuclinoItem> {
        try {
            const response = await this.doThrottledGetRequest(`/items/${itemId}`);
            switch (response.data.object) {
                case 'collection': {
                    return NuclinoCollection.fromRawData(response.data);

                }
                case 'item': {
                    return NuclinoArticle.fromRawData(response.data);

                }

            }
        } catch (error) {
            console.error(
                `Error fetching content for item ${itemId}:`,
                error.message
            );
            throw error;
        }
    }


    /* public async getFile(fileId: string): Promise<NuclinoFile> {
         try {
             const response = await this.doThrottledGetRequest(`/files/${fileId}`);
             return NuclinoFile.fromRawData(response.data.data);
         } catch (error) {
             console.error(
                 `Error fetching file ${fileId}:`,
                 error.message
             );
             throw error;
         }
     }*/


    public async getAllItems(workspaceId, after) {
        try {
            const response = await this.doThrottledGetRequest(`/items`, {
                params: {workspaceId, limit: 100, after},
            });
            return response
        } catch (error) {
            console.error('Error fetching items from Nuclino:', error);
            throw error;
        }
    }


}

/**
 * And this is for all the dust.
 */

class Dust {
    constructor(
        private rateLimiter: Bottleneck,
        private workspaceId: string,
    ) {
        this.rateLimiter = rateLimiter;
        this.workspaceId = workspaceId;
    }

    public async upsertArticleToDustDatasource(article: NuclinoArticle, author: NuclinoUser, lastUpdater: NuclinoUser, breadCrumb: string, destination: Dustination) {
        const fullPath = `${breadCrumb}`
        const content = `
Path: ${fullPath}
Url: ${article.url}
Title: ${article.title}
Author: ${author.firstName} ${author.lastName}
Last updated by: ${lastUpdater.firstName} ${lastUpdater.lastName}
Created At: ${article.createdAt}
Updated At: ${article.lastUpdatedAt}
Content:
${article.content}
  `.trim();

        try {
            await this.doThrottledPostRequest(
                `/vaults/${destination.spaceId}/data_sources/${destination.sourceId}/documents/${article.id}`,
                {
                    title: article.title,
                    mime_type: 'text/markdown',
                    text: content,
                    source_url: article.url,
                }
            );
        } catch (error) {
            console.error(`Error upserting article ${article.id} to Dust datasource:`, error);
        }
    }



    private async doThrottledGetRequest(url: string, data?: any) {
        const params = new URLSearchParams();
        if (data?.params) {
            Object.entries(data.params).forEach(([key, value]) => {
                if (typeof value === "string" || typeof value === "number") {
                    params.append(key, "" + value);
                }
            });
        }

        const queryString = params.toString();
        const fullUrl = `https://dust.tt/api/v1/w/${this.workspaceId}${url}${queryString ? `?${queryString}` : ''}`;

        return this.rateLimiter.schedule(async () => {
            try {
            const response = await fetch(fullUrl, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${DUST_API_KEY}`,
                    'Accept': 'application/json'
                },
                signal: AbortSignal.timeout(60 * 1000), // <=== HERE
            });

            if (!response.ok) {
                throw new Error(`HTTP error! Status: ${response.status}`);
            }

            return await response.json();
            } catch (error) {
                console.log("error", error);
                throw error;
            }
        });
    }

    public async doThrottledPostRequest(url: string, data?: any, config?: any) {
        const fullUrl = `https://dust.tt/api/v1/w/${this.workspaceId}${url}`

        const response = await fetch(fullUrl, {
            method: 'POST',
            headers: {
                Authorization: `Bearer ${DUST_API_KEY}`,
                'Content-Type': 'application/json',
            },
            // Include the body as JSON
            body: JSON.stringify(data),
            signal: AbortSignal.timeout(60 * 1000), // <=== HERE
            // Add any other fetch options from config
            ...config
        });

        if (!response.ok) {
            throw new Error(`HTTP error! Status: ${response.status}`);
        }

        return await response.json();
    }


    private async doThrottledDeleteRequest(url: string, data?: any) {
        const params = new URLSearchParams();
        if (data?.params) {
            Object.entries(data.params).forEach(([key, value]) => {
                if (typeof value === "string" || typeof value === "number") {
                    params.append(key, "" + value);
                }
            });
        }

        const queryString = params.toString();
        const fullUrl = `https://dust.tt/api/v1/w/${this.workspaceId}${url}${queryString ? `?${queryString}` : ''}`;

        return this.rateLimiter.schedule(async () => {
            const response = await fetch(fullUrl, {
                method: 'DELETE',
                headers: {
                    'Authorization': `Bearer ${DUST_API_KEY}`,
                    'Accept': 'application/json'
                },
                signal: AbortSignal.timeout(60 * 1000), // <=== HERE
            });

            if (!response.ok) {
                throw new Error(`HTTP error! Status: ${response.status}`);
            }

            return await response.json();
        });
    }




    public async getItems(destination: Dustination, limit?: number, offset?: number): Promise<{
        documents: any,
        total: number
    }> {
        try {
            const results = await this.doThrottledGetRequest(
                `/spaces/${destination.spaceId}/data_sources/${destination.sourceId}/documents`,
                {
                    params: {
                        limit: limit ? limit : 100,
                        offset: offset ? offset : 0
                    }
                })
            ;
            return results;
        } catch (error) {
            console.error('Error fetching items from Dust:', error);
        }
    }

    public async removeDocument(destination: Dustination, documentId: string) {
        try {
            return this.doThrottledDeleteRequest(`/vaults/${destination.spaceId}/data_sources/${destination.sourceId}/documents/${documentId}`);
        } catch (error) {
            console.error('Error removing document from Dust:', error);
        }
    }
}

/**
 * Ha, A little wordplay
 */
interface Dustination {
    spaceId: string,
    sourceId: string;
}


class Mode {
    private constructor(private readonly value: string) {
    }


    static fromString(value: string): Mode {
        if (["sync", "archive", "dryrun"].indexOf(value) === -1) {
            throw new Error(`Invalid Mode: ${value}`);

        }
        return new Mode(value);
    }

    toString(): string {
        return this.value;
    }

    equals(other: Mode): boolean {
        return this.value === other.value;
    }

    static dryrun(): Mode {
        return new Mode('dryrun');
    }

    static archive(): Mode {
        return new Mode('archive');
    }

    static sync(): Mode {
        return new Mode('sync');
    }

}


class NuclinoSyncJob {

    private dustArticles: Map<string, number> = new Map<string, number>()

    constructor(
        private nuclino: Nuclino,
        private dust: Dust,
        private workspaceName: string,
        private destination: Dustination
    ) {
    }

    public async run(mode: Mode) {
        /**
         * The strategy is as follows:
         * 1. Get all documents in Dust data source and mark the last update date
         * 2. Recurse into nuclino workspace and upsert all documents with a last update date after the one we found in dust
         * 3. Any dust documents that were not found in nuclino will be removed from dust
         */
        const workspace = await nuclino.getWorkspace(this.workspaceName);

        await this.fetchAllDustArticles();
        await this.recursiveSync(workspace, workspace.name, mode);
        if (mode.equals(Mode.sync())) {
            await this.removeDustArticlesNotInNuclino();
        }
    }


    private async recursiveSync(nuclinoItem: NuclinoItem, breadCrumb: string, mode: Mode) {
        switch (nuclinoItem.object) {
            case 'item': {
                if (nuclinoItem instanceof NuclinoArticle) {
                    if (!this.dustArticles.has(nuclinoItem.id) || this.dustArticles.get(nuclinoItem.id) < nuclinoItem.lastUpdatedAt.getTime()) {
                        const author = await this.nuclino.getUserInfo(nuclinoItem.createdUserId);
                        const lastUpdater = await this.nuclino.getUserInfo(nuclinoItem.lastUpdatedUserId);
                        if (!mode.equals(Mode.dryrun())) {
                            await this.dust.upsertArticleToDustDatasource(nuclinoItem, author, lastUpdater, breadCrumb, this.destination);
                        }
                        console.log(`Upserted ${nuclinoItem.id}`);
                    } else {
                        console.log(`Skipping ${nuclinoItem.id} because it was updated before the last update date`);
                    }
                    this.dustArticles.delete(nuclinoItem.id);
                }
                return;
            }
            case 'workspace':
            case 'collection': {
                for (const item of nuclinoItem.childIds) {
                    const itemData = await this.nuclino.getItem(item);
                    await this.recursiveSync(itemData, `${breadCrumb}/${itemData.title}`, mode);

                }
            }

        }

    }

    private async fetchAllDustArticles(): void {
        let response = {total: 0};
        let offset = 0;
        const limit = 200;
        do {
            response = await this.dust.getItems(this.destination, limit, offset);



            for (const document of response.documents) {
                this.dustArticles.set(document.document_id, document.timestamp);
            }
            offset += limit;

        } while (response.total > offset);
        console.log(this.dustArticles.size + " articles found in dust");

    }

    private async removeDustArticlesNotInNuclino(): Promise<void> {
        for (const articleId of this.dustArticles.keys()) {
            console.log(`Removing ${articleId} from Dust`);
            await this.dust.removeDocument(this.destination, articleId);
        }
    }

}



/**
 * Bootstrap
 */

// Create a Bottleneck limiter for Dust API
const dustLimiter = new Bottleneck({
    minTime: 500, // 500ms between requests. limit of 120 upserts / minute per workspace https://docs.dust.tt/reference/rate-limits
    maxConcurrent: 1, // Only 1 request at a time
});

const nuclinoLimiter = new Bottleneck({
    minTime: 500, // 400 between requests. per https://help.nuclino.com/b147124e-rate-limiting
    maxConcurrent: 1, // Only 1 request at a time
});



/**
 * Time to run some code
 */

dotenv.config();

// Configuration
const {
    NUCLINO_API_KEY,
    DUST_API_KEY,
    DUST_WORKSPACE_ID,
} = process.env;


const missingEnvVars = [
    ['NUCLINO_API_KEY', NUCLINO_API_KEY],
    ['DUST_API_KEY', DUST_API_KEY],
    ['DUST_WORKSPACE_ID', DUST_WORKSPACE_ID],
].filter(([name, value]) => !value).map(([name]) => name);

if (missingEnvVars.length > 0) {
    throw new Error(`Please provide values for the following environment variables in the .env file: ${missingEnvVars.join(', ')}`);
}


const nuclino = new Nuclino(nuclinoLimiter);

const dust = new Dust(dustLimiter, DUST_WORKSPACE_ID);



// Run the migration
if (process.argv.length < 6 || ["sync", "archive", "dryrun"].indexOf(process.argv[5])) {
    console.error('Usage: npm run import <nuclinoWorkspaceName> <DustSpaceId> <DustDataSourceId> <mode>sync|archive|dryrun</mode>');
    process.exit(1);
}
const job = new NuclinoSyncJob(nuclino, dust, process.argv[2], {spaceId: process.argv[3], sourceId: process.argv[4]});
job.run(Mode.fromString(process.argv[5]));
