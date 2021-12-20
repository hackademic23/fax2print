const { Client } = require('yapople');
const { print, getPrinters } = require('pdf-to-printer');
const fs = require('fs');

const config = {
	debug: false,
	timer: 60000,
	options: {
		printer: 'Brother MFC-9340CDW Printer',
		scale: 'fit'
	},
	dir: './pdfs/'
};

const client = new Client({
	host: 'outlook.office365.com',
	port: 995,
	tls: true,
	mailparser: true,
	username: 'user@example.com\\faxes@example.com',
	password: ''
});

if(config.debug) {
	getPrinters().then(console.log);
}

(async function run() {
	if(fs.existsSync(config.dir)) {
		fs.rmSync(config.dir, { recursive: true });
	}
	fs.mkdirSync(config.dir);

	await client.connect();
	
	console.log(`Checking for messages...`);
	const messages = await client.retrieveAndDeleteAll();

	if(messages.length == 0) {
		console.log(`No messages yet...`);
	} else {
		console.log(`We have messages, processing now...`);
	}

	messages.forEach((message, index, array) => {
		if(message.attachments) {
			message.attachments.forEach((attachment, index, array) => {

				// Check the content type is PDF.
				if(attachment.contentType == 'application/pdf') {
					// Generate a file a name.
					var attachmentName = attachment.fileName.toString().replace('.pdf', '_' + index + '.pdf');

					// Announce the processing.
					console.log(`Processing PDF attachment ${attachmentName}`);

					// Write the contents to the file system.
					fs.writeFileSync(config.dir + attachmentName, attachment.content);

					// Print the file from the file system.
					print(config.dir + attachmentName, config.options);
				}
			});
		}
	});

	await client.quit();

	setTimeout(run, config.timer);
})().catch(console.error);