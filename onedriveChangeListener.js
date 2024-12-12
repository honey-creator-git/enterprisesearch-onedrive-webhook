const cron = require("node-cron");
const axios = require("axios");
const client = require("./elasticsearch");
require("dotenv").config();

// Function to get access token from Microsoft Graph API
const getAccessToken = async (tenantId, clientId, clientSecret) => {
    const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append("client_id", clientId);
    params.append("scope", "https://graph.microsoft.com/.default");
    params.append("grant_type", "client_credentials");
    params.append("client_secret", clientSecret);

    try {
        const response = await axios.post(tokenEndpoint, params, {
            headers: {
                "Content-Type": "application/x-www-form-urlencoded",
            },
        });
        return response.data.access_token;
    } catch (error) {
        console.error("Error getting access token:", error.message);
        throw new Error("Error getting access token");
    }
};

async function recreateOneDriveSubscription(accessToken, userName) {
    console.log("Re-creating OneDrive subscription for user:", userName);
    const subscriptionData = {
        changeType: "updated", // Listen for all updates (add, update, delete)
        notificationUrl: process.env.NOTIFICATION_URL,  // URL to receive notifications
        resource: `users/${userName}/drive/root`,  // Listen for changes in the user's OneDrive
        expirationDateTime: new Date(Date.now() + 60 * 60 * 1000).toISOString(),  // 1 hour from now
        clientState: "clientStateValue"  // Optional: any state to track the subscription
    };

    const graphBaseUrl = "https://graph.microsoft.com/v1.0";

    try {
        const response = await axios.post(
            `${graphBaseUrl}/subscriptions`,
            subscriptionData,
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    "Content-Type": "application/json"
                }
            }
        );
        console.log("Subscription re-created:", response.data);  // Log the response for debugging
        return response.data;
    } catch (error) {
        if (error.response) {
            console.error("Error response:", error.response.data);  // Log full error response
            console.error("Error status:", error.response.status);  // Log status code
        } else {
            console.error("Error message:", error.message);  // Log message if response is unavailable
        }
        throw new Error("Failed to re-create subscription");
    }
}

// Function to fetch stored credentials from Elasticsearch
const getStoredCredentials = async () => {
    try {
        const indicesResponse = await client.cat.indices({ format: "json" });
        const indices = indicesResponse
            .map((index) => index.index)
            .filter((name) => name.startsWith("datasource_onedrive_connection_"));

        console.log("Found indices: ", indices);

        let expiredConfigs = [];

        for (const index of indices) {
            const query = {
                query: {
                    range: {
                        expirationDateTime: {
                            lt: "now", // Filter for expired expirationDateTime
                        },
                    },
                },
            };

            const result = await client.search({
                index,
                body: query,
            });

            result.hits.hits.forEach((doc) => {
                const { tenantId, clientId, clientSecret, userName, expirationDateTime } = doc._source;
                expiredConfigs.push({
                    tenantId,
                    clientId,
                    clientSecret,
                    userName,
                    expirationDateTime,
                    docId: doc._id, // Save the document ID for future updates
                    indexName: index, // Keep track of the index name
                });
            });
        }

        return expiredConfigs;
    } catch (error) {
        console.error("Error fetching stored credentials from Elasticsearch:", error.message);
        throw new Error("Failed to fetch stored credentials");
    }
};

// Function to update Elasticsearch with the new expiration date
const updateExpirationDateInElasticsearch = async (indexName, docId, newExpirationDateTime) => {
    try {
        await client.update({
            index: indexName,
            id: docId,
            body: {
                doc: {
                    expirationDateTime: newExpirationDateTime,
                    updatedAt: new Date().toISOString(),
                },
            },
        });
        console.log(`Updated expirationDateTime for docId ${docId} in index ${indexName}`);
    } catch (error) {
        console.error(`Error updating expirationDateTime in Elasticsearch for docId ${docId}:`, error.message);
        throw new Error("Failed to update expirationDateTime in Elasticsearch");
    }
};

// Main function to listen for expired OneDrive subscriptions and recreate them
exports.oneDriveChangeListener = async () => {
    try {
        console.log("Starting OneDrive subscription re-creation job...");

        // Step 1: Get expired OneDrive credentials from Elasticsearch
        const expiredConfigs = await getStoredCredentials();

        // Step 2: Recreate OneDrive subscription for each expired configuration
        for (const config of expiredConfigs) {
            const { tenantId, clientId, clientSecret, userName, docId, indexName } = config;

            // Step 3: Get the access token for the OneDrive subscription
            const accessToken = await getAccessToken(tenantId, clientId, clientSecret);

            // Step 4: Recreate the OneDrive subscription
            await recreateOneDriveSubscription(accessToken, userName);
            console.log(`OneDrive subscription recreated for user: ${userName}`);

            // Step 5: Update Elasticsearch with the new expirationDateTime
            const newExpirationDateTime = new Date(Date.now() + 3600000).toISOString(); // New expiration time (1 hour from now)
            await updateExpirationDateInElasticsearch(indexName, docId, newExpirationDateTime);
        }

        console.log("OneDrive subscription re-creation job executed successfully.");
    } catch (error) {
        console.error("Error executing OneDrive subscription re-creation job:", error.message);
    }
};
