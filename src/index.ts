import * as dotenv from 'dotenv';
import { sp } from '@pnp/sp';
import { SPFetchClient } from '@pnp/nodejs';
import * as GraphClient from '@microsoft/microsoft-graph-client';

dotenv.config({ path: '.env.development' });

sp.setup({
    sp: {
        fetchClientFactory: () => {
            return new SPFetchClient(
                process.env.SPO_SITE_URL,
                process.env.SPO_CLIENT_ID,
                process.env.SPO_CLIENT_SECRET
            );
        },
        baseUrl: process.env.SPO_SITE_URL
    }
});

(async () => {
    try {
        const client: GraphClient.Client = GraphClient.Client.init({
            authProvider: done => done(null, 'MyAccessToken')
        });

        const result = await sp.search({
            SelectProperties: ['Title', 'Path'],
            Querytext: 'test',
            RowLimit: 500
        });
        console.log('Search Result: ', result);
    } catch (error) {
        console.log('Error: ', JSON.stringify(error, null, 2));
    }
})();

// test();
