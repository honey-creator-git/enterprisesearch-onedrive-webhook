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

// Function to recreate OneDrive subscription
const recreateOneDriveSubscription = async (accessToken, userName) => {
    try {
        const response = await axios.post(
            "https://graph.microsoft.com/v1.0/subscriptions",
            {
                "changeType": "created,updated,deleted",
                "notificationUrl": process.env.NOTIFICATION_URL, // The URL for receiving webhooks
                "resource": `/users/${userName}/drive/root`,
                "expirationDateTime": new Date(Date.now() + 3600000).toISOString(), // Expires in 1 hour
                "clientState": "secretClientValue",
            },
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    "Content-Type": "application/json",
                },
            }
        );

        return response.data;
    } catch (error) {
        console.error("Error creating OneDrive subscription:", error.message);
        throw new Error("Error creating OneDrive subscription");
    }
};

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
