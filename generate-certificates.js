const { google } = require('googleapis');
require('dotenv').config();
const fs = require('fs');
const path = require('path');
const archiver = require('archiver');
const nodemailer = require('nodemailer');

// Parse command line arguments
function parseArgs() {
  const args = process.argv.slice(2);
  const params = {};
  
  for (let i = 0; i < args.length; i += 2) {
    const key = args[i].replace('--', '');
    const value = args[i + 1];
    params[key] = value;
  }
  
  return params;
}

const params = parseArgs();
const { sheetId, sheetName, date, todate, school } = params;

if (!sheetId || !sheetName || !date || !todate) {
  console.error('Missing required parameters. Usage:');
  console.error('node generate-certificates.js --sheet-id <ID> --sheet-name <NAME> --date <YYYY-MM-DD> --todate <YYYY-MM-DD> [--school <SCHOOL>]');
  process.exit(1);
}

// Handle optional school parameter
const schoolFilter = school && school !== 'undefined' ? school : null;

// Google Sheets and Drive credentials
const oauth2Client = new google.auth.OAuth2(
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET,
  process.env.REDIRECT_URI
);

oauth2Client.setCredentials({ refresh_token: process.env.REFRESH_TOKEN });

const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
const drive = google.drive({ version: 'v3', auth: oauth2Client });
const slides = google.slides({ version: 'v1', auth: oauth2Client });

// Nodemailer configuration
const transporter = nodemailer.createTransporter({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL,
    pass: process.env.PASSWORD
  }
});

async function getSheetData(sheetId, sheetName) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: sheetName,
  });
  return response.data.values;
}

// Helper function to format the date
function formatDateToReadable(date) {
  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  const day = date.getDate();
  const year = date.getFullYear();
  const month = monthNames[date.getMonth()];

  let suffix = "th";
  if (day === 1 || day === 21 || day === 31) {
    suffix = "st";
  } else if (day === 2 || day === 22) {
    suffix = "nd";
  } else if (day === 3 || day === 23) {
    suffix = "rd";
  }

  return `${day}${suffix} ${month}, ${year}`;
}

async function sendCertificates(sheetData, date, todate) {
  const templateId = process.env.TEMPLATE_ID;
  const folderId = process.env.FOLDER_ID;

  console.log(`Starting to send ${sheetData.length - 1} certificates via email...`);

  for (let i = 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    const name = row[0]?.toString() || '';
    const email = row[1]?.toString() || '';
    const schoolName = row[2]?.toString() || '';
    const domain = row[3]?.toString() || '';
    const certificateNumber = row[4]?.toString().toUpperCase() || '';
    const formattedDate = formatDateToReadable(new Date(date));
    const formattedtoDate = formatDateToReadable(new Date(todate));

    console.log(`Processing certificate ${i}/${sheetData.length - 1}: ${name}`);

    const copyFile = await drive.files.copy({
      fileId: templateId,
      requestBody: {
        name: `${name} - Certificate`,
        parents: [folderId]
      }
    });

    const copyId = copyFile.data.id;

    function capitalizeWords(domain) {
      return domain.replace(/\b\w/g, char => char.toUpperCase());
    }

    let formattedWebinarName = capitalizeWords(domain);

    function capitalizesch(schoolName) {
      return schoolName.replace(/\b\w/g, char => char.toUpperCase());
    }

    let formattedsch = capitalizesch(schoolName);

    await slides.presentations.batchUpdate({
      presentationId: copyId,
      requestBody: {
        requests: [
          { replaceAllText: { containsText: { text: '{{Name}}' }, replaceText: name } },
          { replaceAllText: { containsText: { text: '{{SchoolName}}' }, replaceText: formattedsch } },
          { replaceAllText: { containsText: { text: '{{WebinarName}}' }, replaceText: formattedWebinarName } },
          { replaceAllText: { containsText: { text: '{{Date}}' }, replaceText: formattedDate } },
          { replaceAllText: { containsText: { text: '{{Dateto}}' }, replaceText: formattedtoDate } },
          { replaceAllText: { containsText: { text: '{{CERT-NUMBER}}' }, replaceText: certificateNumber } }
        ],
      },
    });

    const response = await drive.files.export({
      fileId: copyId,
      mimeType: 'application/pdf',
    }, { responseType: 'stream' });

    const filename = `${name}_${certificateNumber}.pdf`;
    await sendEmailWithAttachment(
      email,
      `Summer Internship Completion Certificate`,
      `Dear ${name},<br><br>
       Greetings from GeniusHub!!<br><br>
       We sincerely appreciate your participation in the <b>GeniusHub Summer Internship Program - 2025</b>. Your dedication and hard work have been truly commendable.<br><br>
       As a token of our appreciation, please find your <b>Internship Completion Certificate</b> attached. We wish you all the best in your future endeavors.<br><br>
       We are also excited to announce our <b>upcoming Autumn Internship Program</b> starting from October, offering opportunities to learn, grow, and work on exciting projects.<br><br>
       Secure your spot through: <a href="https://forms.gle/XjwGdJPuVydMy7xq7" target="_blank">https://forms.gle/XjwGdJPuVydMy7xq7</a> and continue your journey with us to gain a real-world edge for your portfolio.<br><br>
       Looking forward to welcoming you again.<br><br>
       <b>Best Regards,</b><br><br>
       <b>Nisha Jain</b><br>
       <b>Internships Program Manager</b><br>
       <b>+91-9873331785</b><br>
       <img src="https://upload.wikimedia.org/wikipedia/commons/a/a5/Instagram_icon.png" alt="Instagram" width="20" height="20" style="vertical-align:middle;"> 
       <a href="https://www.instagram.com/geniushub_internships" target="_blank">@geniushub_internships</a><br>
       <a href="https://www.geniushub.in/" target="_blank">https://www.geniushub.in</a>`,
      response.data,
      filename
    );

    await drive.files.update({
      fileId: copyId,
      requestBody: { trashed: true }
    });

    await new Promise(resolve => setTimeout(resolve, 2000)); // Reduced delay for GitHub Actions
  }
}

// Function to generate certificates as a single PDF
async function generateCertificatesAsSinglePDF(sheetData, date, todate) {
  const templateId = process.env.TEMPLATE_ID;
  const folderId = process.env.FOLDER_ID;

  console.log(`Starting to generate single PDF with ${sheetData.length - 1} certificates...`);

  // Dynamically import PDFMerger
  const { default: PDFMerger } = await import('pdf-merger-js');
  const merger = new PDFMerger();

  for (let i = 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    const name = row[0]?.toString() || '';
    const schoolName = row[2]?.toString() || '';
    const domain = row[3]?.toString() || '';
    const certificateNumber = row[4]?.toString().toUpperCase() || '';
    const formattedDate = formatDateToReadable(new Date(date));
    const formattedtoDate = formatDateToReadable(new Date(todate));

    console.log(`Processing certificate ${i}/${sheetData.length - 1}: ${name}`);

    const copyFile = await drive.files.copy({
      fileId: templateId,
      requestBody: {
        name: `${name} - Certificate`,
        parents: [folderId]
      }
    });

    const copyId = copyFile.data.id;

    function capitalizeWords(domain) {
      return domain.replace(/\b\w/g, char => char.toUpperCase());
    }

    let formattedWebinarName = capitalizeWords(domain);

    function capitalizesch(schoolName) {
      return schoolName.replace(/\b\w/g, char => char.toUpperCase());
    }

    let formattedsch = capitalizesch(schoolName);

    await slides.presentations.batchUpdate({
      presentationId: copyId,
      requestBody: {
        requests: [
          { replaceAllText: { containsText: { text: '{{Name}}' }, replaceText: name } },
          { replaceAllText: { containsText: { text: '{{SchoolName}}' }, replaceText: formattedsch } },
          { replaceAllText: { containsText: { text: '{{WebinarName}}' }, replaceText: formattedWebinarName } },
          { replaceAllText: { containsText: { text: '{{Date}}' }, replaceText: formattedDate } },
          { replaceAllText: { containsText: { text: '{{Dateto}}' }, replaceText: formattedtoDate } },
          { replaceAllText: { containsText: { text: '{{CERT-NUMBER}}' }, replaceText: certificateNumber } }
        ],
      },
    });

    const response = await drive.files.export({
      fileId: copyId,
      mimeType: 'application/pdf',
    }, { responseType: 'stream' });

    const filePath = path.join(__dirname, `temp_${name}_${certificateNumber}.pdf`);
    const writeStream = fs.createWriteStream(filePath);
    response.data.pipe(writeStream);

    await new Promise(resolve => writeStream.on('finish', resolve));

    try {
      await merger.add(filePath);
      fs.unlinkSync(filePath);
    } catch (err) {
      console.error(`Failed to process file ${filePath}:`, err);
    }

    try {
      await drive.files.update({
        fileId: copyId,
        requestBody: { trashed: true }
      });
    } catch (err) {
      console.error(`Failed to update file ${copyId} status:`, err);
    }
  }

  await merger.save(path.join(__dirname, 'certificates.pdf'));
  console.log('Single PDF generated successfully: certificates.pdf');
}

// Function to generate certificates as a ZIP file
async function generateCertificates(sheetData, date, todate) {
  if (!Array.isArray(sheetData) || sheetData.length === 0) {
    throw new Error('No data found in the Google Sheet.');
  }

  const templateId = process.env.TEMPLATE_ID;
  const folderId = process.env.FOLDER_ID;
  let generatedCount = 0;

  console.log(`Starting to generate ZIP with ${sheetData.length - 1} certificates...`);

  const zipFilePath = path.join(__dirname, 'certificates.zip');
  const output = fs.createWriteStream(zipFilePath);
  const archive = archiver('zip', { zlib: { level: 9 } });

  output.on('close', () => {
    console.log(`Created zip file with ${archive.pointer()} total bytes.`);
  });

  archive.on('error', (err) => {
    throw err;
  });

  archive.pipe(output);

  const totalCertificates = sheetData.length;
  
  for (let i = 1; i < totalCertificates; i++) {
    const row = sheetData[i];
    const name = row[0]?.toString() || '';
    const schoolName = row[2]?.toString() || '';
    const domain = row[3]?.toString() || '';
    const certificateNumber = row[4]?.toString().toUpperCase() || '';

    if (!name || !schoolName || !certificateNumber) {
      console.log(`Skipping row ${i + 1} due to missing data.`);
      continue;
    }

    const formattedDate = formatDateToReadable(new Date(date));
    const formattedtoDate = formatDateToReadable(new Date(todate));

    console.log(`Processing certificate ${i}/${totalCertificates - 1}: ${name} - ${certificateNumber}`);

    try {
      const copyFile = await drive.files.copy({
        fileId: templateId,
        requestBody: {
          name: `${name} - Certificate`,
          parents: [folderId]
        }
      });

      const copyId = copyFile.data.id;

      function capitalizeWord(domain) {
        return domain.replace(/\b\w/g, char => char.toUpperCase());
      }
  
      let formattedWebinar = capitalizeWord(domain);

      function capitalizeschool(schoolName) {
        return schoolName.replace(/\b\w/g, char => char.toUpperCase());
      }
  
      let formattedschool = capitalizeschool(schoolName);

      await slides.presentations.batchUpdate({
        presentationId: copyId,
        requestBody: {
          requests: [
            { replaceAllText: { containsText: { text: '{{Name}}' }, replaceText: name } },
            { replaceAllText: { containsText: { text: '{{SchoolName}}' }, replaceText: formattedschool } },
            { replaceAllText: { containsText: { text: '{{WebinarName}}' }, replaceText: formattedWebinar } },
            { replaceAllText: { containsText: { text: '{{Date}}' }, replaceText: formattedDate } },
            { replaceAllText: { containsText: { text: '{{Dateto}}' }, replaceText: formattedtoDate } },
            { replaceAllText: { containsText: { text: '{{CERT-NUMBER}}' }, replaceText: certificateNumber } }
          ],
        },
      });

      const response = await drive.files.export({
        fileId: copyId,
        mimeType: 'application/pdf',
      }, { responseType: 'stream' });

      const filePath = path.join(__dirname, `${name}_${certificateNumber}.pdf`);
      const writeStream = fs.createWriteStream(filePath);

      response.data.pipe(writeStream);

      await new Promise((resolve, reject) => {
        writeStream.on('finish', resolve);
        writeStream.on('error', reject);
      });

      archive.file(filePath, { name: `${name}_${certificateNumber}.pdf` });

      try {
        await drive.files.update({
          fileId: copyId,
          requestBody: { trashed: true }
        });
      } catch (err) {
        console.error(`Failed to update file ${copyId} status:`, err);
      }

      try {
        fs.unlinkSync(filePath);
      } catch (err) {
        console.error(`Failed to delete file ${filePath}:`, err);
      }

      generatedCount++;

    } catch (err) {
      console.error(`Error processing certificate for ${name} - ${certificateNumber}:`, err);
    }
  }

  console.log('Finalizing ZIP archive');
  await archive.finalize();
  console.log('ZIP file generated successfully: certificates.zip');
}

async function sendEmailWithAttachment(to, subject, htmlContent, pdfStream, filename) {
  const mailOptions = {
    from: '"GeniusHub - Unleash your Genius" <Certificate@geniushub.in>',
    to,
    subject,
    html: htmlContent,
    attachments: [
      {
        filename,
        content: pdfStream,
        contentType: 'application/pdf'
      }
    ]
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log(`Email sent successfully to ${to}`);
  } catch (error) {
    console.error(`Error sending email to ${to}:`, error);
  }
}

// Main execution
async function main() {
  try {
    console.log('Starting certificate generation...');
    console.log(`Sheet ID: ${sheetId}`);
    console.log(`Sheet Name: ${sheetName}`);
    console.log(`Date: ${date} to ${todate}`);
    console.log(`School: ${schoolFilter || 'All schools'}`);

    const sheetData = await getSheetData(sheetId, sheetName);
    const filteredData = schoolFilter ? sheetData.filter(row => row[2]?.toString().toUpperCase() === schoolFilter.toUpperCase()) : sheetData;
    const totalCertificates = filteredData.length - 1;

    console.log(`Found ${totalCertificates} certificates to generate`);

    if (!schoolFilter) {
      // Generate all certificates and send via email
      await sendCertificates(filteredData, date, todate);
      await new Promise(resolve => setTimeout(resolve, 2000));
      await generateCertificatesAsSinglePDF(filteredData, date, todate);
    } else {
      // Specific school selected, generate certificates as a ZIP file
      await generateCertificates(filteredData, date, todate);
    }

    console.log('Certificate generation completed successfully!');
  } catch (error) {
    console.error('Error in certificate generation:', error);
    process.exit(1);
  }
}

main();
