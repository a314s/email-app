const express = require('express');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const nodemailer = require('nodemailer');
const multer = require('multer');
const cors = require('cors');
const dotenv = require('dotenv');

// Load environment variables
dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static('public'));

// Configure multer for file uploads
const upload = multer({
  dest: 'uploads/',
  limits: { fileSize: 10 * 1024 * 1024 } // 10MB limit
});

// Google Sheets API setup
async function getAuthClient() {
  const auth = new google.auth.GoogleAuth({
    keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly']
  });
  return auth.getClient();
}

// Get Google Sheets data
app.get('/api/sheets/:spreadsheetId', async (req, res) => {
  try {
    const { spreadsheetId } = req.params;
    const authClient = await getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    
    // Get all sheets in the spreadsheet
    const sheetList = await sheets.spreadsheets.get({
      spreadsheetId
    });
    
    const sheetNames = sheetList.data.sheets.map(sheet => sheet.properties.title);
    
    res.json({ success: true, sheets: sheetNames });
  } catch (error) {
    console.error('Error fetching sheets:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Get data from a specific sheet
app.get('/api/sheets/:spreadsheetId/:sheetName', async (req, res) => {
  try {
    const { spreadsheetId, sheetName } = req.params;
    const authClient = await getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: sheetName
    });
    
    const rows = response.data.values;
    
    if (!rows || rows.length === 0) {
      return res.status(404).json({ success: false, error: 'No data found' });
    }
    
    // Process the data to group by company
    const headers = rows[0];
    const data = rows.slice(1).map(row => {
      const rowData = {};
      headers.forEach((header, index) => {
        rowData[header] = row[index] || '';
      });
      return rowData;
    });
    
    // Group data by company (assuming company is in column A)
    const groupedData = [];
    let currentCompany = '';
    let companyData = [];
    
    data.forEach(row => {
      const company = row[headers[0]]; // Company name from column A
      
      if (company && company !== currentCompany) {
        if (currentCompany) {
          groupedData.push({
            company: currentCompany,
            contacts: companyData
          });
        }
        currentCompany = company;
        companyData = [row];
      } else {
        companyData.push(row);
      }
    });
    
    // Add the last company group
    if (currentCompany && companyData.length > 0) {
      groupedData.push({
        company: currentCompany,
        contacts: companyData
      });
    }
    
    res.json({ 
      success: true, 
      headers, 
      data: groupedData
    });
  } catch (error) {
    console.error('Error fetching sheet data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Get contact by line number
app.get('/api/contact/line/:spreadsheetId/:sheetName/:lineNumber', async (req, res) => {
  try {
    const { spreadsheetId, sheetName, lineNumber } = req.params;
    const line = parseInt(lineNumber);
    
    if (isNaN(line) || line < 2) { // Line 1 is headers, so contacts start at line 2
      return res.status(400).json({ success: false, error: 'Invalid line number' });
    }
    
    const authClient = await getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!${line}:${line}`
    });
    
    const row = response.data.values?.[0];
    
    if (!row) {
      return res.status(404).json({ success: false, error: 'No data found at this line' });
    }
    
    // Get headers to map column names
    const headerResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!1:1`
    });
    
    const headers = headerResponse.data.values?.[0] || [];
    
    // Map row data to headers
    const contact = {};
    headers.forEach((header, index) => {
      contact[header] = row[index] || '';
    });
    
    res.json({ success: true, contact });
  } catch (error) {
    console.error('Error fetching contact by line:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Get contact by name
app.get('/api/contact/name/:spreadsheetId/:sheetName/:name', async (req, res) => {
  try {
    const { spreadsheetId, sheetName, name } = req.params;
    const authClient = await getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    
    // Get all data from the sheet
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: sheetName
    });
    
    const rows = response.data.values;
    
    if (!rows || rows.length === 0) {
      return res.status(404).json({ success: false, error: 'No data found' });
    }
    
    const headers = rows[0];
    const nameColumnIndex = headers.findIndex(header => header.toLowerCase() === 'name');
    
    if (nameColumnIndex === -1) {
      return res.status(404).json({ success: false, error: 'Name column not found' });
    }
    
    // Find the row with the matching name
    let contactRow = null;
    let rowIndex = -1;
    
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (row[nameColumnIndex] && row[nameColumnIndex].toLowerCase() === name.toLowerCase()) {
        contactRow = row;
        rowIndex = i + 1; // +1 because Google Sheets is 1-indexed
        break;
      }
    }
    
    if (!contactRow) {
      return res.status(404).json({ success: false, error: 'Contact not found' });
    }
    
    // Map row data to headers
    const contact = {};
    headers.forEach((header, index) => {
      contact[header] = contactRow[index] || '';
    });
    
    // Add the line number
    contact.lineNumber = rowIndex;
    
    res.json({ success: true, contact });
  } catch (error) {
    console.error('Error fetching contact by name:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Upload Word template
app.post('/api/template/upload', upload.single('template'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ success: false, error: 'No file uploaded' });
  }
  
  const templatePath = path.join(__dirname, 'templates', req.file.originalname);
  
  // Ensure templates directory exists
  if (!fs.existsSync(path.join(__dirname, 'templates'))) {
    fs.mkdirSync(path.join(__dirname, 'templates'), { recursive: true });
  }
  
  // Move file from uploads to templates directory
  fs.copyFileSync(req.file.path, templatePath);
  fs.unlinkSync(req.file.path); // Remove from uploads
  
  res.json({ 
    success: true, 
    message: 'Template uploaded successfully',
    filename: req.file.originalname
  });
});

// List available templates
app.get('/api/templates', (req, res) => {
  const templatesDir = path.join(__dirname, 'templates');
  
  // Create templates directory if it doesn't exist
  if (!fs.existsSync(templatesDir)) {
    fs.mkdirSync(templatesDir, { recursive: true });
    return res.json({ success: true, templates: [] });
  }
  
  const templates = fs.readdirSync(templatesDir)
    .filter(file => file.endsWith('.docx'))
    .map(file => ({
      filename: file,
      path: path.join(templatesDir, file)
    }));
  
  res.json({ success: true, templates });
});

// Generate email content from template
app.post('/api/generate-email', (req, res) => {
  try {
    const { templateName, contactData } = req.body;
    
    if (!templateName || !contactData) {
      return res.status(400).json({ success: false, error: 'Missing required parameters' });
    }
    
    const templatePath = path.join(__dirname, 'templates', templateName);
    
    if (!fs.existsSync(templatePath)) {
      return res.status(404).json({ success: false, error: 'Template not found' });
    }
    
    // Read the template
    const content = fs.readFileSync(templatePath, 'binary');
    const zip = new PizZip(content);
    
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true
    });
    
    // Set the template data
    doc.setData(contactData);
    
    // Render the document
    doc.render();
    
    // Get the rendered content
    const buffer = doc.getZip().generate({ type: 'nodebuffer' });
    
    // Convert to HTML (simplified approach - in a real app, you might want to use a more robust docx-to-html converter)
    // For this example, we'll just extract the text content
    const textContent = buffer.toString('utf8').replace(/[^\x20-\x7E]/g, ' ');
    
    res.json({ 
      success: true, 
      emailContent: textContent,
      subject: `Information for ${contactData.Name || 'you'}`,
      to: contactData.Email || ''
    });
  } catch (error) {
    console.error('Error generating email:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Send email
app.post('/api/send-email', (req, res) => {
  try {
    const { to, subject, content, from } = req.body;
    
    if (!to || !subject || !content) {
      return res.status(400).json({ success: false, error: 'Missing required parameters' });
    }
    
    // Configure email transport
    const transporter = nodemailer.createTransport({
      host: process.env.SMTP_HOST,
      port: process.env.SMTP_PORT,
      secure: process.env.SMTP_SECURE === 'true',
      auth: {
        user: process.env.SMTP_USER,
        pass: process.env.SMTP_PASS
      }
    });
    
    // Send email
    transporter.sendMail({
      from: from || process.env.DEFAULT_FROM_EMAIL,
      to,
      subject,
      html: content
    }, (error, info) => {
      if (error) {
        console.error('Error sending email:', error);
        return res.status(500).json({ success: false, error: error.message });
      }
      
      res.json({ 
        success: true, 
        message: 'Email sent successfully',
        messageId: info.messageId
      });
    });
  } catch (error) {
    console.error('Error in send-email endpoint:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Open in Outlook
app.post('/api/open-outlook', (req, res) => {
  try {
    const { to, subject, body } = req.body;
    
    if (!to) {
      return res.status(400).json({ success: false, error: 'Recipient email is required' });
    }
    
    // Generate mailto URL
    const mailtoUrl = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject || '')}&body=${encodeURIComponent(body || '')}`;
    
    res.json({ 
      success: true, 
      mailtoUrl,
      message: 'Use this URL to open in Outlook'
    });
  } catch (error) {
    console.error('Error generating mailto URL:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`Open http://localhost:${PORT} in your browser`);
});
