const fs = require("fs");
const createCsvWriter = require("csv-writer").createObjectCsvWriter;
const createCsvArrayWriter = require("csv-writer").createArrayCsvWriter;
const moment = require("moment");
const path = require("path");

async function generateCloudCaptureReports() {
  if (!fs.existsSync("./attachments")) {
    fs.mkdirSync("./attachments");
  }

  if (!fs.existsSync("./inputFiles")) {
    fs.mkdirSync("./inputFiles");
  }

  // Set the directory path
  const currentDate = moment().format("YYYY-MM-DD");
  const currentDateTime = moment().format("YYYY-MM-DD HH:mm");
  //const directoryPath = `./attachments/gen2_${currentDate}`;
  const directoryPath = `./attachments/cloudCaptureReport`;

  if (!fs.existsSync(directoryPath)) {
    fs.mkdirSync(directoryPath);
  }

  // Set the CSV output file path
  const csvFilePath = `./inputFiles/cloud-capture-report.csv`;

  if (!fs.existsSync("./inputFiles")) {
    fs.mkdirSync("./inputFiles");
  }

  // Define CSV headers
  const csvHeaders = [
    "Date",
    "Network",
    "Number of messages ingested",
    "Number of messages exported",
    "Number of messages not exported",
    "Number of backlog messages exported in the last 30 days",
    "Number of messages not exported in the last 30 days",
    "Email Sent Date",
  ];

  // Define CSV writer
  const csvArrayWriter = createCsvArrayWriter({
    header: csvHeaders,
    path: csvFilePath,
    append: true, // append to existing file if it exists
  });

  // Check if file exists
  if (!fs.existsSync(csvFilePath)) {
    // Create CSV file with headers if it doesn't exist
    csvArrayWriter
      .writeRecords([csvHeaders])
      .then(() => {
        console.log("CSV file created successfully");
        console.log("-------------------------");
      })
      .catch((err) => {
        console.error(err);
      });
  } else {
    console.log("CSV file exists.");
    console.log("-------------------------");
  }

  // Create a CSV writer object
  const csvWriter = createCsvWriter({
    path: csvFilePath,
    append: true, // append to existing file if it exists
    header: [
      { id: "Date", title: "Date" },
      { id: "Network", title: "Network" },
      {
        id: "Number of messages ingested",
        title: "Number of messages ingested",
      },
      {
        id: "Number of messages exported",
        title: "Number of messages exported",
      },
      {
        id: "Number of messages not exported",
        title: "Number of messages not exported",
      },
      {
        id: "Number of backlog messages exported in the last 30 days",
        title: "Number of backlog messages exported in the last 30 days",
      },
      {
        id: "Number of messages not exported in the last 30 days",
        title: "Number of messages not exported in the last 30 days",
      },
      { id: "Email Sent Date", title: "Email Sent Date" },
    ],
  });

  // Read the files in the directory
  fs.readdir(directoryPath, async (err, files) => {
    if (err) {
      console.log("Error reading directory:", err);
      return;
    }
    // Loop through each file in the directory
    files.map(async (file) => {
      // Get the full path of the file
      const filePath = `${directoryPath}/${file}`;
      const fileExtension = path.extname(filePath);

      if (fileExtension !== ".xlsx" && fileExtension !== ".xls") {
        // Read the content of the file
        fs.readFile(filePath, "utf8", async (err, emailText) => {
          if (err) {
            console.log("Error reading file:", err);
            return;
          }

          const subjectDateRegex = /Cloud\s*Capture\s*Reconciliation\s*Report\s*-\s*(\d{2}\/\d{2}\/\d{4})/;
          const match = subjectDateRegex.exec(emailText);
          let sent_date = match ? match[1] : null;

          // Split the email text by lines
          const lines = emailText.split("\n");

          // Initialize an array to store the extracted data
          const extractedData = [];

          // Define a flag to indicate when to start processing data
          let isDataSection = false;

          const regexPattern =
            /^(\d{2}\/\d{2}\/\d{4})\s+([A-Za-z\s]+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)$/;

          const headerRegexPattern = /Date\s+(Network)\s+(Number of messages ingested)\s+(Number of messages exported)\s+(Number of messages not exported)\s+(Number of backlog messages exported in the last 30 days)\s+(Number of messages not exported in the last 30 days)/;


          for (const line of lines) {

            if (headerRegexPattern.test(line.trim())) {
              // Extract and count spaces between column headers
                isDataSection = true;
                continue;
            }

            if (isDataSection) {
              // const data = line.trim().split(/\s+/);
              // if (data.length === 7) {
              const jsonData = line.trim().match(regexPattern);
              if (jsonData) {
                const [_, date, network, ingested, exported, notExported, NumberOfBacklogMessagesExportedLast30Days, NumberOfBacklogMessagesNotExportedLast30Days] =
                  jsonData;
                extractedData.push({
                  Date: date,
                  Network: network,
                  "Number of messages ingested": parseInt(ingested),
                  "Number of messages exported": parseInt(exported),
                  "Number of messages not exported": parseInt(notExported),
                  "Number of backlog messages exported in the last 30 days": parseInt(NumberOfBacklogMessagesExportedLast30Days),
                  "Number of messages not exported in the last 30 days": parseInt(NumberOfBacklogMessagesNotExportedLast30Days),
                  "Email Sent Date": sent_date,
                });
              } else {
                // Stop processing data when there are not enough columns
                isDataSection = false;
              }
            }
          }

          if (extractedData.length > 0) {
            await csvWriter.writeRecords(extractedData);
            console.log("cloudCaptureReport Data Generated!");
          }

          fs.unlink(filePath, (error) => {
              if (error) {
                  console.log(`Error Deleting cloudCaptureReport (${file}) File!`);
              } else {
                  console.log(`cloudCaptureReport (${file}) File Deleted Successfully!`);
              }
          })
        });
      }
    });
  });
}

module.exports = {
  generateCloudCaptureReports,
};
