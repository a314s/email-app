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
const OpenAI = require('openai');
const database = require('./database');

// Load environment variables
dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

// Initialize database
database.initializeDatabase()
    .then(() => {
        console.log('Database initialized successfully');
    })
    .catch(err => {
        console.error('Error initializing database:', err);
    });

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
    const contactData = req.body.contactData || {};
    
    if (!to || !subject || !content) {
      return res.status(400).json({ success: false, error: 'Missing required parameters' });
    }
    
    // Create a transporter
    const transporter = nodemailer.createTransport({
      host: 'smtp.gmail.com',
      port: 587,
      secure: false,
      auth: {
        user: from || process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASSWORD
      }
    });
    
    // Send the email
    const info = await transporter.sendMail({
      from: from || process.env.EMAIL_USER,
      to,
      subject,
      html: content
    });
    
    // Record the email in the database
    const emailId = await database.recordEmail({
      recipient: to,
      subject,
      content,
      contactData
    });
    
    res.json({
      success: true,
      message: 'Email sent successfully',
      emailId,
      messageId: info.messageId
    });
  } catch (error) {
    console.error('Error sending email:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Get calendar data for a specific month
app.get('/api/calendar/:year/:month', async (req, res) => {
    try {
        const { year, month } = req.params;
        const calendarData = await database.getCalendarData(parseInt(year), parseInt(month));
        res.json({ success: true, calendarData: calendarData.calendarData });
    } catch (error) {
        console.error('Error fetching calendar data:', error);
        res.status(500).json({ success: false, error: error.message });
    }
});

// Improve email content using OpenAI ChatGPT API
app.post('/api/improve-email', async (req, res) => {
  try {
    const { content, contactData } = req.body;
    
    if (!content) {
      return res.status(400).json({ success: false, error: 'Email content is required' });
    }
    
    // Check if OpenAI API key is available
    const apiKey = process.env.OPENAI_API_KEY;
    if (!apiKey) {
      return res.status(400).json({ success: false, error: 'OpenAI API key is not configured' });
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
    
    console.log('Calling OpenAI API with key:', apiKey.substring(0, 5) + '...');
    
    try {
      // Initialize the OpenAI client
      const openai = new OpenAI({
        apiKey: apiKey,
      });
      
      // Prepare the user message
      const userMessage = `You are an expert email writer for a company called Navitech Aid. I need you to improve the following email content.
Make it more professional, industry-specific, and persuasive, but maintain the same general tone and intent while keeping it concise and to the point.
Make sure the output is the same language as the original email.
Do not add any new information that isn't implied in the original email.
Do not change any factual information.
Focus on improving the language, structure, and persuasiveness.

${contactContext}

Original Email:
${content}

Improved Email:`;

      console.log('Sending message to OpenAI...');
      
      // Call OpenAI API
      const response = await openai.chat.completions.create({
        model: "gpt-4",
        messages: [
          {
            role: "user",
            content: userMessage
          }
        ],
        max_tokens: 4000,
        temperature: 0.7
      });
      
      console.log('OpenAI API response received');
      
      // Extract the improved content from the response
      const improvedContent = response.choices[0].message.content;
      
      res.json({
        success: true,
        improvedContent
      });
    } catch (apiError) {
      console.error('Error calling OpenAI API:', apiError);
      return res.status(500).json({
        success: false,
        error: 'Failed to improve email content',
        details: apiError.message
      });
    }
  } catch (error) {
    console.error('Error improving email content:', error);
    res.status(500).json({ success: false, error: 'Failed to improve email content' });
  }
});

// Update Excel sheet with new data
app.post('/api/excel/update', (req, res) => {
  try {
    const { filename, sheetName, formData } = req.body;
    
    if (!filename || !sheetName || !formData) {
      return res.status(400).json({ success: false, error: 'Missing required parameters' });
    }
    
    const filePath = path.join(__dirname, 'excel', filename);
    
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ success: false, error: 'File not found' });
    }
    
    // Read the Excel file
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
    
    // Create a new row with the form data
    const newRow = [];
    headers.forEach(header => {
      newRow.push(formData[header] || '');
    });
    
    // Add the new row to the data
    jsonData.push(newRow);
    
    // Convert back to worksheet
    const newWorksheet = XLSX.utils.aoa_to_sheet(jsonData);
    
    // Update the workbook
    workbook.Sheets[sheetName] = newWorksheet;
    
    // Write the updated workbook back to the file
    XLSX.writeFile(workbook, filePath);
    
    res.json({ 
      success: true, 
      message: 'Data added successfully',
      rowNumber: jsonData.length
    });
  } catch (error) {
    console.error('Error updating Excel file:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Calendar and Follow-up API Endpoints

// Record a sent email
app.post('/api/emails/record', (req, res) => {
  try {
    const { recipient, subject, content, contactData } = req.body;
    
    if (!recipient || !subject || !content) {
      return res.status(400).json({ success: false, error: 'Missing required fields' });
    }
    
    // First, record the email
    database.recordEmail({ recipient, subject, content, contactData })
      .then(emailId => {
        // Then, check if we have follow-up data from Excel
        if (contactData && contactData.excelFile && contactData.sheetName) {
          database.importFollowUpDataFromExcel(contactData.excelFile, contactData.sheetName, contactData)
            .then(followUpData => {
              if (followUpData) {
                // Update the email with Excel follow-up data
                database.updateEmailWithExcelData(emailId, followUpData)
                  .then(() => {
                    res.json({ success: true, emailId });
                  })
                  .catch(error => {
                    console.error('Error updating email with Excel data:', error);
                    // Still return success as the email was recorded
                    res.json({ success: true, emailId, warning: 'Email recorded but Excel data not updated' });
                  });
              } else {
                res.json({ success: true, emailId });
              }
            })
            .catch(error => {
              console.error('Error importing follow-up data from Excel:', error);
              // Still return success as the email was recorded
              res.json({ success: true, emailId, warning: 'Email recorded but Excel data not imported' });
            });
        } else {
          res.json({ success: true, emailId });
        }
      })
      .catch(error => {
        console.error('Error recording email:', error);
        res.status(500).json({ success: false, error: error.message });
      });
  } catch (error) {
    console.error('Error recording email:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Get all sent emails
app.get('/api/emails', (req, res) => {
  database.getSentEmails()
    .then(emails => {
      res.json({ success: true, emails });
    })
    .catch(error => {
      console.error('Error fetching emails:', error);
      res.status(500).json({ success: false, error: error.message });
    });
});

// Get emails by date
app.get('/api/emails/date/:date', (req, res) => {
  try {
    const { date } = req.params;
    const dateObj = new Date(date);
    
    if (isNaN(dateObj.getTime())) {
      return res.status(400).json({ success: false, error: 'Invalid date format' });
    }
    
    database.getEmailsByDate(dateObj)
      .then(emails => {
        res.json({ success: true, emails });
      })
      .catch(error => {
        console.error('Error fetching emails by date:', error);
        res.status(500).json({ success: false, error: error.message });
      });
  } catch (error) {
    console.error('Error fetching emails by date:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Get all pending follow-ups
app.get('/api/follow-ups', (req, res) => {
  database.getPendingFollowUps()
    .then(followUps => {
      res.json({ success: true, followUps });
    })
    .catch(error => {
      console.error('Error fetching follow-ups:', error);
      res.status(500).json({ success: false, error: error.message });
    });
});

// Get follow-ups by date
app.get('/api/follow-ups/date/:date', (req, res) => {
  try {
    const { date } = req.params;
    const dateObj = new Date(date);
    
    if (isNaN(dateObj.getTime())) {
      return res.status(400).json({ success: false, error: 'Invalid date format' });
    }
    
    database.getFollowUpsByDate(dateObj)
      .then(followUps => {
        res.json({ success: true, followUps });
      })
      .catch(error => {
        console.error('Error fetching follow-ups by date:', error);
        res.status(500).json({ success: false, error: error.message });
      });
  } catch (error) {
    console.error('Error fetching follow-ups by date:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Complete a follow-up
app.post('/api/follow-ups/:id/complete', (req, res) => {
  try {
    const { id } = req.params;
    const { notes } = req.body;
    
    database.completeFollowUp(id, notes)
      .then(result => {
        res.json({ 
          success: true, 
          message: 'Follow-up completed successfully',
          ...result
        });
      })
      .catch(error => {
        console.error('Error completing follow-up:', error);
        res.status(500).json({ success: false, error: error.message });
      });
  } catch (error) {
    console.error('Error completing follow-up:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Get all companies
app.get('/api/companies', (req, res) => {
  try {
    database.getAllCompanies()
      .then(companies => {
        res.json({ success: true, companies });
      })
      .catch(error => {
        console.error('Error fetching companies:', error);
        res.status(500).json({ success: false, error: error.message });
      });
  } catch (error) {
    console.error('Error fetching companies:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Get calendar data for a specific company
app.get('/api/calendar/:year/:month/:company', (req, res) => {
  try {
    const { year, month, company } = req.params;
    const yearNum = parseInt(year);
    const monthNum = parseInt(month);
    
    if (isNaN(yearNum) || isNaN(monthNum) || monthNum < 1 || monthNum > 12) {
      return res.status(400).json({ success: false, error: 'Invalid year or month' });
    }
    
    if (!company) {
      return res.status(400).json({ success: false, error: 'Company name is required' });
    }
    
    database.getCompanyCalendarData(yearNum, monthNum, decodeURIComponent(company))
      .then(result => {
        res.json(result);
      })
      .catch(error => {
        console.error('Error fetching company calendar data:', error);
        res.status(500).json({ success: false, error: error.message });
      });
  } catch (error) {
    console.error('Error fetching company calendar data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Start the server
const server = app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`Open http://localhost:${PORT} in your browser`);
  
  // Initialize the database
  database.initializeDatabase()
    .then(() => {
      console.log('Database initialized');
    })
    .catch(err => {
      console.error('Error initializing database:', err);
    });
});
