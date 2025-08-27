// worker.js
/**
 * This is a complete Cloudflare Worker script that acts as a flexible
 * external options provider for Lark Approval.
 *
 * It connects to your Lark Base and retrieves data from a specified field,
 * formatting it according to the Lark API documentation.
 *
 * How to deploy this worker:
 * 1. Go to your Cloudflare dashboard and navigate to the "Workers & Pages" section.
 * 2. Create a new Worker.
 * 3. Copy and paste the entire code from this file into the Worker's script editor.
 * 4. In the Worker's "Settings" tab, go to "Variables" and add the following
 * environment variables:
 *
 * - LARK_APP_ID: Your Lark app ID.
 * - LARK_APP_SECRET: Your Lark app secret.
 * - LARK_AUTH_TOKEN: A secret token you define for authentication (e.g., "my-secret-token-123").
 *
 * 5. Save and deploy the Worker.
 * 6. Use the Worker's public URL in your Lark Approval form's external options settings.
 */

// Global variable to cache the Lark access token
let appAccessToken = null;
let tokenExpiry = 0;

/**
 * Fetches and caches the Lark app access token.
 * @param {object} env The environment variables.
 * @returns {Promise<string>} The access token.
 */
async function getAppAccessToken(env) {
    const currentTime = Math.floor(Date.now() / 1000);
    // Refresh token if it's older than 3000 seconds (50 minutes) or not set
    if (!appAccessToken || (currentTime - tokenExpiry) > 3000) {
        const url = "https://open.larksuite.com/open-apis/auth/v3/app_access_token/internal";
        const headers = { "Content-Type": "application/json; charset=utf-8" };
        const data = {
            "app_id": env.LARK_APP_ID,
            "app_secret": env.LARK_APP_SECRET
        };

        try {
            const response = await fetch(url, {
                method: 'POST',
                headers: headers,
                body: JSON.stringify(data)
            });
            const result = await response.json();
            if (result.code !== 0) {
                throw new Error(result.msg || 'Failed to get app access token.');
            }
            appAccessToken = result["app_access_token"];
            tokenExpiry = currentTime;
        } catch (error) {
            console.error('Error fetching app access token:', error);
            throw new Error('Failed to get app access token.');
        }
    }
    return appAccessToken;
}

/**
 * Fetches records from a specific Lark Base table.
 * @param {string} appToken The Base ID.
 * @param {string} tableId The table ID.
 * @param {string} fieldName The name of the field to retrieve.
 * @param {object} env The environment variables.
 * @returns {Promise<Array>} An array of records.
 */
async function getRecords(appToken, tableId, fieldName, env) {
    const accessToken = await getAppAccessToken(env);
    const url = `https://open.larksuite.com/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records`;

    const headers = {
        "Authorization": `Bearer ${accessToken}`,
        "Content-Type": "application/json"
    };
    const params = new URLSearchParams({
        page_size: '100',
        field_names: JSON.stringify([fieldName])
    });

    try {
        const response = await fetch(`${url}?${params.toString()}`, {
            method: 'GET',
            headers: headers
        });
        const result = await response.json();
        if (result.code !== 0 || !result.data || !result.data.items) {
            console.warn('No items found or unexpected response:', result);
            return [];
        }
        return result.data.items;
    } catch (error) {
        console.error('Error fetching records from Lark Base:', error);
        throw new Error('Failed to get records from Lark Base.');
    }
}

/**
 * Main handler for all incoming requests to the Worker.
 * @param {Request} request The incoming request.
 * @param {object} env The environment variables configured in Cloudflare.
 * @returns {Response} The response to be sent back.
 */
async function handleRequest(request, env) {
    // Handle GET requests for a simple health check
    if (request.method === 'GET') {
        return new Response('Lark Base External Options API is running!', { status: 200 });
    }

    // Only allow POST requests for the API endpoint
    if (request.method !== 'POST') {
        return new Response('Method not allowed', { status: 405 });
    }

    try {
        // Parse the JSON request body
        const { app_token, table_id, field_name, token } = await request.json();

        // Validate the secret token
        if (token !== env.LARK_AUTH_TOKEN) {
            return new Response(JSON.stringify({
                code: 1,
                msg: "Invalid token.",
                data: {}
            }), { status: 401, headers: { 'Content-Type': 'application/json' } });
        }

        // Validate required parameters
        if (!app_token || !table_id || !field_name) {
            return new Response(JSON.stringify({
                code: 1,
                msg: "Missing required parameters: app_token, table_id, or field_name.",
                data: {}
            }), { status: 400, headers: { 'Content-Type': 'application/json' } });
        }

        const records = await getRecords(app_token, table_id, field_name, env);

        const options = [];
        const i18nResources = [];
        const uniqueValues = new Set();

        records.forEach(record => {
            const fieldValue = record.fields[field_name];
            
            if (fieldValue === null || fieldValue === undefined || fieldValue === '') {
                return;
            }

            let valueToDisplay;
            let optionsId = record.id;
            
            if (Array.isArray(fieldValue)) {
                valueToDisplay = fieldValue.map(item => {
                    if (item && item.name) {
                        return item.name;
                    }
                    return item;
                }).join(', ');
            } else if (typeof fieldValue === 'object' && fieldValue !== null && fieldValue.name) {
                valueToDisplay = fieldValue.name;
            } else {
                valueToDisplay = fieldValue.toString();
            }

            if (!uniqueValues.has(valueToDisplay)) {
                options.push({
                    id: optionsId,
                    value: valueToDisplay,
                });
                uniqueValues.add(valueToDisplay);
            }
        });

        const result = {
            code: 0,
            msg: "success",
            data: {
                result: {
                    options: options,
                    i18nResources: i18nResources,
                    hasMore: false,
                    nextPageToken: ""
                }
            }
        };

        return new Response(JSON.stringify(result), {
            status: 200,
            headers: { 'Content-Type': 'application/json' }
        });

    } catch (e) {
        console.error('Error handling request:', e.message);
        return new Response(JSON.stringify({
            code: 1,
            msg: "Internal server error. Check worker logs.",
            data: {}
        }), { status: 500, headers: { 'Content-Type': 'application/json' } });
    }
}

// Add the event listener to respond to all incoming requests
addEventListener('fetch', event => {
    event.respondWith(handleRequest(event.request, event.env));
});
