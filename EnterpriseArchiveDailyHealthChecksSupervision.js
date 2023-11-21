const fs = require("fs");
const createCsvWriter = require("csv-writer").createObjectCsvWriter;
const createCsvArrayWriter = require("csv-writer").createArrayCsvWriter;
const moment = require("moment");
const path = require("path");

async function generateArchiveDailyHealthChecksSupervisionReport() {
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
  const directoryPath = `./attachments/enterpriseArchiveDailyHealthChecksSupervision`;

  if (!fs.existsSync(directoryPath)) {
    fs.mkdirSync(directoryPath);
  }

  // Set the CSV output file path
  const csvFilePath = `./inputFiles/enterprise-archive-daily-health-checks-supervision.csv`;

  if (!fs.existsSync("./inputFiles")) {
    fs.mkdirSync("./inputFiles");
  }

  // Define CSV headers
  const csvHeaders = [
    "Queue Name",
    "Queue ID",
    "Number of Policies",
    "Scheduled Time",
    "Number of Runs",
    "Number of Items Processed into the Queue",
    "Echo Close",
    "Echo Update",
    "Start Time",
    "End Time",
    'Exported Date',
    'Sent Date'
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
      { id: "Queue Name", title: "Queue Name" },
      { id: "Queue ID", title: "Queue ID" },
      {
        id: "Number of Policies",
        title: "Number of Policies",
      },
      {
        id: "Scheduled Time",
        title: "Scheduled Time",
      },
      {
        id: "Number of Runs",
        title: "Number of Runs",
      },
      {
        id: "Number of Items Processed into the Queue",
        title: "Number of Items Processed into the Queue",
      },
      {
        id: "Echo Close",
        title: "Echo Close",
      },
      { id: "Echo Update", title: "Echo Update" },
      { id: "Start Time", title: "Start Time" },
      { id: "End Time", title: "End Time" },
      { id: "Exported Date", title: "Exported Date" },
      { id: "Sent Date", title: "Sent Date" },
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

          const subjectDateRegex = /Enterprise Archive Daily Health Checks - Supervision - From (\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2})/;
          const subjectMatch = subjectDateRegex.exec(emailText);
          let sent_date = subjectMatch ? subjectMatch[1] : null;

          // Define a regular expression pattern to capture the data under the specified headers
          const regexPattern =
            /(\S.*?)\s+(\S+)\s+(\d+)\s+(\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}(?:\.\d+)?)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}\.\d+)(?:.*?(\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}\.\d+))?/g;

          // Initialize an array to store the extracted data
          const extractedData = [];

          // Match the pattern in the email text and extract the data
          let match;
          while ((match = regexPattern.exec(emailText)) !== null) {
            const [
              _,
              queueName,
              queueId,
              numberOfPolicies,
              scheduledTime,
              numberOfRuns,
              numberOfItemsProcessed,
              echoClose,
              echoUpdate,
              startTime,
              endTime,
            ] = match;

            // Create an object with the extracted data, handling empty fields
            const dataObject = {
              "Queue Name": queueName,
              "Queue ID": queueId,
              "Number of Policies": parseInt(numberOfPolicies) || 0,
              "Scheduled Time": scheduledTime,
              "Number of Runs": parseInt(numberOfRuns) || 0,
              "Number of Items Processed into the Queue":
                parseInt(numberOfItemsProcessed) || 0,
              "Echo Close": parseInt(echoClose) || 0,
              "Echo Update": parseInt(echoUpdate) || 0,
              "Start Time": startTime || null,
              "End Time": endTime || null, // Set endTime to null if it's not captured
              "Exported Date": moment().format('YYYY-MM-DD HH:mm'), // Set endTime to null if it's not captured
              "Sent Date": sent_date, // Set endTime to null if it's not captured
            };

            // Add the object to the array
            extractedData.push(dataObject);
          }

          if (extractedData.length > 0) {
            await csvWriter.writeRecords(extractedData);
            console.log("enterpriseArchiveDailyHealthChecksSupervisionReport Data Generated!");
          }

          fs.unlink(filePath, (error) => {
            if (error) {
              console.log(`Error Deleting enterpriseArchiveDailyHealthChecksSupervisionReport (${file}) File!`);
            } else {
              console.log(
                `enterpriseArchiveDailyHealthChecksSupervisionReport (${file}) File Deleted Successfully!`
              );
            }
          });

        });
      }
    });
  });
}

module.exports = {
  generateArchiveDailyHealthChecksSupervisionReport
};
