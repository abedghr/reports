const fs = require('fs');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const createCsvArrayWriter = require('csv-writer').createArrayCsvWriter;
const moment = require('moment');
const path = require('path');

async function generateArchived() {
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
    const csvFilePath = `./inputFiles/health-checks-archived-reports.csv`;

    if (!fs.existsSync('./inputFiles')) {
        fs.mkdirSync('./inputFiles');
    }

    // Define CSV headers
    const csvHeaders = [
        'Content Source', 'Network Archived', 'Count', 'Exported Date', 'Email Sent Date'
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
            { id: 'Content Source', title: 'Content Source' },
            { id: 'Network Archived', title: 'Network Archived' },
            { id: 'Count', title: 'Count' },
            { id: 'Exported Date', title: 'Exported Date' },
            { id: 'Email Sent Date', title: 'Email Sent Date' }, 
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

                    const regex = /Archived Count([\s\S]*?)Failed Count/g;
                    const matches = regex.exec(data);
                    const extractedData = matches[1].trim();

                    const jsonData = extractedData
                        .split('\n')
                        .filter(line => line.trim() !== '')
                        .map(line => {
                            const [contentSource, networkArchivedCount, Count] = line.split(/\t/);
                            return { 
                                'Content Source': contentSource,
                                'Network Archived': networkArchivedCount,
                                'Count': Count,
                                'Exported Date': currentDateTime,
                                'Email Sent Date': sent_date
                            };
                        });

                    await csvWriter.writeRecords(jsonData);
                    console.log("healthChecksArchived Data Generated!");
                    // fs.unlink(filePath, (error) => {
                    //     if (error) {
                    //         console.log(`Error Deleting Gen2 (${file}) File!`);
                    //     } else {
                    //         console.log(`Gen2 (${file}) File Deleted Successfully!`);
                    //     }
                    // })
                });
            }
        });
    });

};

module.exports = {
    generateArchived
}