const cron = require("node-cron");
const oneDriveChangeListener = require("./onedriveChangeListener").oneDriveChangeListener;

// Schedule the cron job to run every 5 minutes
cron.schedule("*/5 * * * *", async () => {
    try {
        console.log("Executing OneDrive Change listener job...");
        await oneDriveChangeListener();
        console.log("OneDrive Change listener job executed successfully.");
    } catch (error) {
        console.error("Error executing OneDrive Change ID listener job:", error.message);
    }
});

console.log("OneDrive Change Listener is running...");
