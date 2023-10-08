const Imap = require('imap');
const fs = require('fs');
const {simpleParser} = require("mailparser");
const  inspect = require('util').inspect;
const moment = require('moment');
require('dotenv').config();


// IMAP server details
// const imap = new Imap({
//     user: process.env.IMAP_GMAIL_USER,
//     password: process.env.IMAP_GMAIL_PASSWORD,
//     host: process.env.IMAP_GMAIL_HOST,
//     port: process.env.IMAP_PORT,
//     timeout: 10000,
//     tls: {
//         rejectUnauthorized: false
//     }
// });

const imap = new Imap({
    user: process.env.IMAP_OUTLOOK_USER,
    password: process.env.IMAP_OUTLOOK_PASSWORD,
    host: process.env.IMAP_OUTLOOK_HOST,
    port: process.env.IMAP_PORT,
    timeout: 10000,
    tls: {
        rejectUnauthorized: false
    }
});

if (!fs.existsSync('./attachments')) {
    fs.mkdirSync('./attachments');
}


function openInbox(cb) {
    imap.openBox('INBOX', true, cb);
}

function printEmailDetails (prefix, parsed) {
    console.log(prefix + 'date: ' + parsed.date);
    console.log(prefix + 'From: ' + (parsed.from ? parsed.from.text : 'Unknown'));
    console.log(prefix + 'Subject: ' + parsed.subject);
    console.log(prefix + 'Attachments: ' + parsed.attachments.length);
    console.log(prefix + 'TextHtml: ' + parsed.textAsHtml);
    console.log("###################")
    if (parsed.attachments.length > 0) {
        exportAttachments(parsed.attachments);
    }

    if (parsed.textAsHtml) {
        // Do some process on html text
    }
}

function exportAttachments(attachments) {

    const directory = './attachments';

    let attachmentsDir = `${directory}/${moment().format('YYYY-MM-DD')}`;
    if (!fs.existsSync(attachmentsDir)) {
        fs.mkdirSync(attachmentsDir);
    }
    attachments.forEach(function (attachment, index) {
        let fileName = `${attachmentsDir}/${index}_${attachment.filename.replace(/\s+/g, '_')}`
        if (fileName) {
            fs.writeFileSync(fileName, attachment.content);
            console.log(`# ${fileName} has been exported.`);
        }
    });
}


imap.once('ready', function() {
    openInbox(function(err, box) {
        if (err) throw err;

        const date = moment().format("MMM D, YYYY");
        imap.search([ 'ALL', ['SINCE', date], ['FROM', process.env.TARGET_EMAIL] ], function(err, results) {
            if (err) throw err;
            let f = imap.fetch(results, { bodies: '' });
            f.on('message', function(msg, seqno) {
                let prefix = '(#' + seqno + ') ';
                msg.on('body', function(stream, info) {
                    simpleParser(stream, (err, parsed) => {
                        if (err) throw err;
                        printEmailDetails(prefix, parsed)
                    });
                });
                msg.once('attributes', function(attrs) {
                    console.log(prefix + 'Attributes: %s', inspect(attrs, false, 8));
                });
                msg.once('end', function() {
                    console.log(prefix + 'Finished');
                });
            });
            f.once('error', function(err) {
                console.log('Fetch error: ' + err);
            });
            f.once('end', function() {
                console.log('Done fetching all messages!');
                imap.end();
            });
        });
    });
})

imap.connect();
