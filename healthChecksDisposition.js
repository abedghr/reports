const fs = require('fs');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const createCsvArrayWriter = require('csv-writer').createArrayCsvWriter;
const moment = require('moment');
const path = require('path');

async function generateDisposition() {
    if (!fs.existsSync('./attachments')) {
        fs.mkdirSync('./attachments');
    }

    if (!fs.existsSync('./inputFiles')) {
        fs.mkdirSync('./inputFiles');
    }


    // Set the directory path
    const currentDate = moment().format('YYYY-MM-DD');
    const currentDateTime = moment().format('YYYY-MM-DD HH:mm')
    //const directoryPath = `./attachments/gen2_${currentDate}`;
    const directoryPath = `./attachments/healthChecks`;

    if (!fs.existsSync(directoryPath)) {
        fs.mkdirSync(directoryPath);
    }

    // Set the CSV output file path
    const csvFilePath = `./inputFiles/health-checks-disposition-reports.csv`;

    if (!fs.existsSync('./inputFiles')) {
        fs.mkdirSync('./inputFiles');
    }

    // Define CSV headers
    const csvHeaders = [
        'Content Source', 'Network', 'Total Number of Disposed', 'Exported Date', 'Email Sent Date'
    ];

    // Define CSV writer
    const csvArrayWriter = createCsvArrayWriter({
        header: csvHeaders,
        path: csvFilePath,
        append: true // append to existing file if it exists
    });

    // Check if file exists
    if (!fs.existsSync(csvFilePath)) {
        // Create CSV file with headers if it doesn't exist
        csvArrayWriter.writeRecords([csvHeaders]).then(() => {
            console.log('CSV file created successfully');
            console.log('-------------------------');
        }).catch((err) => {
            console.error(err);
        });
    } else {
        console.log('CSV file exists.');
        console.log('-------------------------');
    }

    // Create a CSV writer object
    const csvWriter = createCsvWriter({
        path: csvFilePath,
        append: true, // append to existing file if it exists
        header: [
            { id: 'contentSource', title: 'Content Source' },
            { id: 'network', title: 'Network' },
            { id: 'totalDisposed', title: 'Total Number of Disposed' },
            { id: 'exportedDate', title: 'Exported Date' },
            { id: 'emailSentDate', title: 'Email Sent Date' },
        ]
    });


    // Read the files in the directory
    fs.readdir(directoryPath, async (err, files) => {
        if (err) {
            console.log('Error reading directory:', err);
            return;
        }
        // Loop through each file in the directory
        files.map(async file => {
            // Get the full path of the file
            const filePath = `${directoryPath}/${file}`;
            const fileExtension = path.extname(filePath);

            if (fileExtension !== '.xlsx' && fileExtension !== '.xls') {
                // Read the content of the file
                fs.readFile(filePath, 'utf8', async (err, data) => {
                    if (err) {
                        console.log('Error reading file:', err);
                        return;
                    }


                    const subjectDateRegex = /Daily Health Checks From (\d{4}-\d{2}-\d{2})/;
                    const match = subjectDateRegex.exec(data);
                    let sent_date = match ? match[1] : null;

                    const subjectToDateRegex = /To (\d{4}-\d{2}-\d{2})/;
                    const matchTo = subjectToDateRegex.exec(data);
                    const sent_to_date = matchTo ? matchTo[1] : null;
                    sent_date = `${sent_date} - ${sent_to_date}`;

                    // Define the regular expression pattern to match the data between "Disposition" and "Ingestion Queries"
                    const pattern = /Disposition\s+([\s\S]*?)Ingestion Queries/m;

                    // Find the match in the data
                    const match11 = data.match(pattern);

                    if (match11 && match11[1]) {
                        const dispositionData = match11[1].trim();

                        // Split the paragraph into lines
                        const lines = dispositionData.split('\n');

                        // Remove the first line (header) as we don't need it for the objects
                        lines.shift();

                        // Function to parse each line and create an object
                        function parseLine(line) {
                            let [contentSource, network, unknown, totalDisposed] = line.split(/\s+/);
                            const parseUnknown = parseInt(unknown);

                            network = `${network} ${isNaN(parseUnknown) ? unknown : ''}`;
                            totalDisposed = !isNaN(parseUnknown) ? parseUnknown : parseInt(totalDisposed);
                            return { contentSource, network, totalDisposed, 'exportedDate':  currentDateTime, 'emailSentDate': sent_date };
                        }

                        // Create an array of objects by parsing each line
                        const listOfObjects = lines.map(parseLine);

                        // Output the resulting array of objects
                        await csvWriter.writeRecords(listOfObjects);
                        console.log("healthChecksDisposition Data Generated!");

                    }
                });
            }
        });
    });

};

module.exports = {
    generateDisposition
}
