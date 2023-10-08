const fs = require('fs');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const createCsvArrayWriter = require('csv-writer').createArrayCsvWriter;
const moment = require('moment');
const cheerio = require('cheerio');
const path = require('path');
const XLSX = require('xlsx');

if (!fs.existsSync('./attachments')) {
    fs.mkdirSync('./attachments');
}

if (!fs.existsSync('./inputFiles')) {
    fs.mkdirSync('./inputFiles');
}


// Set the directory path
const currentDate = moment().format('YYYY-MM-DD');
const directoryPath = `./attachments/${currentDate}`;

if (!fs.existsSync(directoryPath)) {
    fs.mkdirSync(directoryPath);
}

// Set the CSV output file path
const csvFilePath = `./inputFiles/reports.csv`;

if (!fs.existsSync('./inputFiles')) {
    fs.mkdirSync('./inputFiles');
}

// Define CSV headers
const csvHeaders = [
    'File Name', 'Repository Name', 'Audit Name', 'Start Date', 'End Date', 'Date Imported', 'Date Of Event', 'Date Completed',
    'Records', 'User Name', 'Doc Class Name', 'Total Msgs', 'Total Size In MB', 'Download Attempts', 'Files Info', 'Status',
    'Import Type', 'Size', 'Raw Hits', 'Unique Hits', 'Number Of Unconvertable Messages',
    'Participants Created', 'Participants Updated', 'Participants Groups Created',

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
        {id: 'file_name', title: 'File Name'},
        {id: 'repository_name', title: 'Repository Name'},
        {id: 'audit_name', title: 'Audit Name'},
        {id: 'start_date', title: 'Start Date'},
        {id: 'end_date', title: 'End Date'},
        {id: 'date_imported', title: 'Date Imported'},
        {id: 'date_of_event', title: 'Date Of Event'},
        {id: 'date_completed', title: 'Date Completed'},
        {id: 'records', title: 'Records'},
        {id: 'user_name', title: 'User Name'},
        {id: 'doc_class_name', title: 'Doc Class Name'},
        {id: 'total_msgs', title: 'Total Msgs'},
        {id: 'total_size_in_mb', title: 'Total Size In MB'},
        {id: 'download_attempts', title: 'Download Attempts'},
        {id: 'files_info', title: 'Files Info'},
        {id: 'status', title: 'Status'},
        {id: 'import_type', title: 'Import Type'},
        {id: 'size', title: 'Size'},
        {id: 'raw_hits', title: 'Raw Hits'},
        {id: 'unique_hits', title: 'Unique Hits'},
        {id: 'number_of_unconvertable_messages', title: 'Number Of Unconvertable Messages'},
        {id: 'participants_created', title: 'Participants Created'},
        {id: 'participants_updated', title: 'Participants Updated'},
        {id: 'participants_groups_created', title: 'Participants Groups Created'}
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
        const matches = [];

        if (fileExtension !== '.xlsx' && fileExtension !== '.xls') {
            // Read the content of the file
            fs.readFile(filePath, 'utf8', async (err, data) => {
                if (err) {
                    console.log('Error reading file:', err);
                    return;
                }
                const regex = /File Name\s*:\s*(.*?)\s*START_DATE\s*=\s*([\d\/: ]+)\s*END_DATE\s*=\s*([\d\/: ]+)\s*Records\s*:\s*(\d+)/gs;

                let match = regex.exec(data);
                while (match !== null) {
                    const file_name = match[1];
                    const start_date = match[2];
                    const end_date = match[3];
                    const records = parseInt(match[4]);

                    matches.push({file_name, start_date, end_date, records});
                    match = regex.exec(data);
                }

                if (matches.length === 0) {
                    // const regex = /From:\s*(\d{4}-\d{2}-\d{2})\s*\n\s*To\s*(\d{4}-\d{2}-\d{2})\s*[\s\S]*?Records:\s*(\d+)/gm;
                    const regex = /From\s*:\s*([\d\-]+)\s*To\s*:*\s*([\d\-]+)\s*[\s\S]*?Records\s*:\s*(\d+)/gm;

                    let match2 = regex.exec(data);
                    while (match2 !== null) {
                        const start_date = match2[1];
                        const end_date = match2[2];
                        const records = parseInt(match2[3]);

                        matches.push({start_date, end_date, records});
                        match2 = regex.exec(data);
                    }

                }

                if (matches.length === 0) {
                    const regex = /From\s*:\s*([\d\-: ]+)\s*To\s*:([\s\S]*?)\n[\s\S]*?Download Summary\s+([\s\S]*?)all successful([\s\S]*([\d,]+)$)/gm;

                    let match3 = regex.exec(data);
                    while (match3 !== null) {
                        const start_date = match3[1].trim();
                        const end_date = match3[2].trim();
                        const download_attempts = parseInt(match3[3]);
                        let files_info = null;
                        if (match3[4] !== undefined) {
                            const files = match3[4].replace(/(-+[\d\S]+)$/gm, '').trim().split('\n');
                            files_info = files.join('*******').replace(/\s+/g, '|');
                        }

                        matches.push({start_date, end_date, download_attempts, files_info});
                        match3 = regex.exec(data)
                    }

                }

                if (matches.length === 0) {
                    // const regex = /From\s*:([\s\S]*?)To\s*:([\s\S]*?)\n[\s\S]*?Download Summary\s+([\s\S]*?)all successful([\s\S]*([\d,]+)$)/gm;
                    const regex = /From\s*:\s*([\d\-: ]+)\s*To\s*:\s*([\s\S]*?)\n[\s\S]*?Download Summary\s+([\s\S]*?)all successful/gm;

                    let match4 = regex.exec(data);
                    while (match4 !== null) {
                        const start_date = match4[1].trim();
                        const end_date = match4[2].trim();
                        const regexSummary = /\s*([\d,]+)\s*FTP/g;
                        const summaryAttempts = match4[3].match(regexSummary);
                        const download_attempts = summaryAttempts.length > 0 ? summaryAttempts[0] : null;
                        matches.push({start_date, end_date, download_attempts});
                        match4 = regex.exec(data);
                    }
                }
                if (matches.length === 0) {
                    const $ = cheerio.load(data);
                    $('*').filter((i, el) => Object.keys(el.attribs).length > 0).removeAttr('id class style');

                    const rows = $('table tbody tr');
                    const result = rows.map((i, row) => {
                        const columns = $(row).find('td');
                        const cellValues = columns.map((j, column) => $(column).text().trim()).get();
                        return {rowNumber: i + 1, cells: cellValues};
                    }).get();

                    if (result.length > 0) {
                        const keys = Object.values(result[0].cells);
                        const outputList = result.map((row, index) => {
                            if (index !== 0) {
                                const obj = {};
                                keys.forEach((key, index) => {
                                    let clearedKey = key.toLowerCase()
                                        .replace(/\s+/g, '_')
                                        .replace('(', '')
                                        .replace(')', '');
                                    obj[clearedKey] = row.cells[index];
                                });
                                obj['rowNumber'] = row.rowNumber;
                                return obj;
                            }
                        });
                        const output = outputList.slice();
                        output.shift();

                        const regex1 = /Total number of searches run: (\d+)/;
                        const regex2 = /Total number of documents returned across all searches: (\d+)/;
                        const numSearches = data.match(regex1);
                        const numDocsReturned = data.match(regex2);
                        if (numSearches && numDocsReturned) {
                            if (output.length > 0) {
                                output.map(record => {
                                    const file_name = record.file_name;
                                    const date_imported = record.date_imported
                                    const status = record.status
                                    const import_type = record.import_type
                                    const size = record.size_kb
                                    const participants_created = record.participants_created
                                    const participants_updated = record.participants_updated
                                    const participants_groups_created = record.participants_groups_created

                                    matches.push({
                                        file_name,
                                        date_imported,
                                        status,
                                        import_type,
                                        size,
                                        participants_created,
                                        participants_updated,
                                        participants_groups_created
                                    });
                                });
                            }
                            console.log("Total number of searches run:", numSearches[1]);
                            console.log("Total number of documents returned across all searches:", numDocsReturned[1]);
                        } else {
                            if (output.length > 0) {
                                output.map(record => {
                                    const repository_name = record.repositoryname;
                                    const doc_class_name = record.docclassname;
                                    const date_of_event = record.dateofevent;
                                    const total_msgs = record.totalmsgs
                                    const total_size_in_mb = record.totalsize_in_mb

                                    matches.push({
                                        repository_name,
                                        doc_class_name,
                                        date_of_event,
                                        total_msgs,
                                        total_size_in_mb,
                                    });
                                });
                            }
                        }
                    }
                }
                // If we have processed all files, write the file details to the CSV file
                if (matches.length > 0) {
                    await csvWriter.writeRecords(matches);
                    console.log('Data written successfully');
                    console.log('-------------------------');
                }
            });
        } else {
            const workbook = XLSX.readFile(filePath);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const result = XLSX.utils.sheet_to_json(worksheet);
            if (result.length > 0) {
                const headersValues = ['Audit Name', 'User Name', 'Date Completed', 'Raw Hits', 'Unique Hits', 'Number Of Unconvertable Messages'];
                let targetIndex = null;

                result.forEach((item, index) => {
                    if (Object.values(item).some(header => headersValues.includes(header))) {
                        targetIndex = index;
                        return true;
                    }
                });

                const headersIndexes = {}
                result.map((row, index) => {
                    let objectValues = Object.values(row);
                    if (index >= targetIndex) {
                        if (index === targetIndex) {
                            objectValues.map((el, index) => {
                                headersIndexes[el.toLowerCase().replace(/\s+/g, '_')] = index
                            })
                        } else {
                            matches.push({
                                audit_name: objectValues[headersIndexes.audit_name],
                                user_name: objectValues[headersIndexes.user_name],
                                date_completed: objectValues[headersIndexes.date_completed],
                                raw_hits: objectValues[headersIndexes.raw_hits],
                                unique_hits: objectValues[headersIndexes.unique_hits],
                                number_of_unconvertable_messages: objectValues[headersIndexes.number_of_unconvertable_messages]
                            });
                        }
                    }
                });
                // If we have processed all files, write the file details to the CSV file
                if (matches.length > 0) {
                    await csvWriter.writeRecords(matches);
                    console.log('Data written successfully');
                    console.log('-------------------------');
                }
            }
        }
    });
});
