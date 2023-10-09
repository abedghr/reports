const cron = require('node-cron');
const { generateStatistics } = require('./detailedArchiveStatistics');
const { generateArchived } = require('./healthChecksArchived');
const { generateDisposition } = require('./healthChecksDisposition');
const { generateFailed } = require('./healthChecksFailedOlder');
const { generateReceived } = require('./healthChecksReceived');
const { generateCloudCaptureReports } = require("./cloudCaptureReport");

// Every minute
// for 4:00 PM => 10 16 * * *
cron.schedule('10 16 * * *', async () => {
    await generateStatistics();
    await generateArchived();
    await generateReceived();
    await generateDisposition();
    await generateFailed();
    await generateCloudCaptureReports();
})