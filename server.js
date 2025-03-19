const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const nodemailer = require('nodemailer');
const multer = require('multer');
const cors = require('cors');
const dotenv = require('dotenv');
const mammoth = require('mammoth');
const axios = require('axios');

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
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    // Determine the destination folder based on file type
    let uploadFolder = 'uploads';
    if (file.fieldname === 'template') {
      uploadFolder = 'templates';
    } else if (file.fieldname === 'excelFile') {
      uploadFolder = 'excel';
    }
    
    // Create the directory if it doesn't exist
    const dir = path.join(__dirname, uploadFolder);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    
    cb(null, uploadFolder);
  },
  filename: function (req, file, cb) {
    // Keep the original filename
    cb(null, file.originalname);
  }
});

const upload = multer({
  storage: storage,
  limits: { fileSize: 10 * 1024 * 1024 } // 10MB limit
});

// Upload Excel file
app.post('/api/excel/upload', upload.single('excelFile'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ success: false, error: 'No file uploaded' });
  }
  
  try {
    const filePath = path.join(__dirname, req.file.destination, req.file.originalname);
    
    // Read the Excel file to get sheet names
    const workbook = XLSX.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    
    res.json({
      success: true,
      message: 'Excel file uploaded successfully',
      filename: req.file.originalname,
      sheets: sheetNames
    });
  } catch (error) {
    console.error('Error processing Excel file:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Get list of uploaded Excel files
app.get('/api/excel/files', (req, res) => {
  const excelDir = path.join(__dirname, 'excel');
  
  // Create excel directory if it doesn't exist
  if (!fs.existsSync(excelDir)) {
    fs.mkdirSync(excelDir, { recursive: true });
    return res.json({ success: true, files: [] });
  }
  
  const files = fs.readdirSync(excelDir)
    .filter(file => file.endsWith('.xlsx') || file.endsWith('.xls'))
    .map(file => ({
      filename: file,
      path: path.join(excelDir, file)
    }));
  
  res.json({ success: true, files });
});

// Get sheet names from Excel file
app.get('/api/excel/:filename/sheets', (req, res) => {
  try {
    const { filename } = req.params;
    const filePath = path.join(__dirname, 'excel', filename);
    
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ success: false, error: 'File not found' });
    }
    
    const workbook = XLSX.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    
    res.json({ success: true, sheets: sheetNames });
  } catch (error) {
    console.error('Error reading Excel file:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Get data from a specific sheet
app.get('/api/excel/:filename/:sheetName', (req, res) => {
  try {
    const { filename, sheetName } = req.params;
    const filePath = path.join(__dirname, 'excel', filename);
    
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ success: false, error: 'File not found' });
    }
    
    const workbook = XLSX.readFile(filePath);
    
    if (!workbook.SheetNames.includes(sheetName)) {
      return res.status(404).json({ success: false, error: 'Sheet not found' });
    }
    
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (jsonData.length === 0) {
      return res.status(404).json({ success: false, error: 'No data found in sheet' });
    }
    
    // Extract headers (first row)
    const headers = jsonData[0];
    
    // Process the data to group by company
    const data = jsonData.slice(1).map(row => {
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
app.get('/api/contact/line/:filename/:sheetName/:lineNumber', (req, res) => {
  try {
    const { filename, sheetName, lineNumber } = req.params;
    const line = parseInt(lineNumber);
    
    if (isNaN(line) || line < 2) { // Line 1 is headers, so contacts start at line 2
      return res.status(400).json({ success: false, error: 'Invalid line number' });
    }
    
    const filePath = path.join(__dirname, 'excel', filename);
    
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ success: false, error: 'File not found' });
    }
    
    const workbook = XLSX.readFile(filePath);
    
    if (!workbook.SheetNames.includes(sheetName)) {
      return res.status(404).json({ success: false, error: 'Sheet not found' });
    }
    
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (jsonData.length < line) {
      return res.status(404).json({ success: false, error: 'Line number out of range' });
    }
    
    // Get headers (first row)
    const headers = jsonData[0];
    
    // Get the row at the specified line number
    const row = jsonData[line - 1]; // -1 because array is 0-indexed
    
    if (!row) {
      return res.status(404).json({ success: false, error: 'No data found at this line' });
    }
    
    // Map row data to headers
    const contact = {};
    headers.forEach((header, index) => {
      contact[header] = row[index] || '';
    });
    
    // Add line number to contact data
    contact.lineNumber = line;
    
    res.json({ success: true, contact });
  } catch (error) {
    console.error('Error fetching contact by line:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Get contact by name
app.get('/api/contact/name/:filename/:sheetName/:name', (req, res) => {
  try {
    const { filename, sheetName, name } = req.params;
    const filePath = path.join(__dirname, 'excel', filename);
    
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ success: false, error: 'File not found' });
    }
    
    const workbook = XLSX.readFile(filePath);
    
    if (!workbook.SheetNames.includes(sheetName)) {
      return res.status(404).json({ success: false, error: 'Sheet not found' });
    }
    
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (jsonData.length === 0) {
      return res.status(404).json({ success: false, error: 'No data found in sheet' });
    }
    
    // Get headers (first row)
    const headers = jsonData[0];
    
    // Find the name column (assumed to be column E, which is index 4)
    const nameColumnIndex = 4; // Column E (0-indexed)
    
    if (nameColumnIndex >= headers.length) {
      return res.status(404).json({ success: false, error: 'Name column not found' });
    }
    
    // Find the row with the matching name - now with partial matching
    let contactRow = null;
    let rowIndex = -1;
    const searchName = name.toLowerCase();
    
    for (let i = 1; i < jsonData.length; i++) {
      const row = jsonData[i];
      if (row[nameColumnIndex] && row[nameColumnIndex].toString().toLowerCase().includes(searchName)) {
        contactRow = row;
        rowIndex = i + 1; // +1 because we want 1-indexed line numbers
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

// Generate email from template
app.post('/api/generate-email', async (req, res) => {
  try {
    const { templateName, contactData } = req.body;
    
    if (!templateName || !contactData) {
      return res.status(400).json({ success: false, error: 'Template name and contact data are required' });
    }
    
    const templatePath = path.join(__dirname, 'templates', templateName);
    
    if (!fs.existsSync(templatePath)) {
      return res.status(404).json({ success: false, error: 'Template not found' });
    }
    
    // Read the template file
    const content = await mammoth.extractRawText({ path: templatePath });
    let emailContent = content.value;
    
    // Replace placeholders with contact data
    for (const [key, value] of Object.entries(contactData)) {
      const placeholder = `{{${key}}}`;
      emailContent = emailContent.replace(new RegExp(placeholder, 'g'), value || '');
    }
    
    // Also replace [Name] with the contact's name
    if (contactData.Name) {
      emailContent = emailContent.replace(/\[Name\]/g, contactData.Name);
    }
    
    // Extract subject from the first line if it starts with "Subject:"
    let subject = '';
    const lines = emailContent.split('\n');
    if (lines[0].startsWith('Subject:')) {
      subject = lines[0].substring('Subject:'.length).trim();
      emailContent = lines.slice(1).join('\n').trim();
    } else {
      // Default subject if not specified in template
      subject = "What can navitech Aid do for you";
    }
    
    res.json({
      success: true,
      emailContent,
      subject
    });
  } catch (error) {
    console.error('Error generating email:', error);
    res.status(500).json({ success: false, error: 'Failed to generate email' });
  }
});

// Send email
app.post('/api/send-email', async (req, res) => {
  try {
    const { to, subject, content, from } = req.body;
    
    if (!to || !subject || !content) {
      return res.status(400).json({ success: false, error: 'Missing required parameters' });
    }
    
    // Configure transporter
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
    await transporter.sendMail({
      from: from || process.env.SMTP_USER,
      to,
      subject,
      text: content
    });
    
    res.json({ success: true, message: 'Email sent successfully' });
  } catch (error) {
    console.error('Error sending email:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Improve email content using Anthropic API
app.post('/api/improve-email', async (req, res) => {
  try {
    const { content, contactData } = req.body;
    
    if (!content) {
      return res.status(400).json({ success: false, error: 'Email content is required' });
    }
    
    // Check if Anthropic API key is available
    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) {
      return res.status(400).json({ success: false, error: 'Anthropic API key is not configured' });
    }
    
    // Prepare contact information for context
    let contactContext = '';
    if (contactData) {
      contactContext = 'Contact Information:\n';
      for (const [key, value] of Object.entries(contactData)) {
        if (value && typeof value === 'string' && value.trim() !== '') {
          contactContext += `${key}: ${value}\n`;
        }
      }
    }
    
    // Prepare the prompt for Anthropic
    const prompt = `
You are an expert email writer for a company called Navitech Aid. I need you to improve the following email content.
Make it more professional, industry-specific, and persuasive, but maintain the same general tone and intent.
Do not add any new information that isn't implied in the original email.
Do not change any factual information.
Focus on improving the language, structure, and persuasiveness.

${contactContext}

Original Email:
${content}

Improved Email:
`;

    try {
      // Call Anthropic API
      const response = await axios.post(
        'https://api.anthropic.com/v1/messages',
        {
          model: 'claude-3-sonnet-20240229',
          max_tokens: 4000,
          messages: [
            {
              role: 'user',
              content: prompt
            }
          ]
        },
        {
          headers: {
            'Content-Type': 'application/json',
            'x-api-key': apiKey,
            'anthropic-version': '2023-06-01'
          }
        }
      );
      
      // Extract the improved content from the response
      const improvedContent = response.data.content[0].text;
      
      res.json({
        success: true,
        improvedContent
      });
    } catch (apiError) {
      console.error('Error calling Anthropic API:', apiError.response?.data || apiError.message);
      return res.status(500).json({ 
        success: false, 
        error: 'Failed to improve email content',
        details: apiError.response?.data || apiError.message
      });
    }
  } catch (error) {
    console.error('Error improving email content:', error);
    res.status(500).json({ success: false, error: 'Failed to improve email content' });
  }
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`Open http://localhost:${PORT} in your browser`);
});
