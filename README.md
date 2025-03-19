# Email Automation App

This application helps you automate the process of sending emails based on data from Google Sheets and Word document templates.

## Features

- Connect to Google Sheets to access contact information
- Search for contacts by line number or name
- Upload and use Word document templates with dynamic content replacement
- Generate emails with personalized content
- Send emails directly or open in Outlook
- View all contacts in a data table
- Get detailed information about each contact

## Setup Instructions

### Prerequisites

- Node.js (v14 or higher)
- Google Sheets API credentials
- Microsoft Word for template creation

### Installation

1. Clone this repository or download the files
2. Open a terminal in the project directory
3. Install dependencies:

```bash
npm install
```

4. Create a `.env` file in the root directory with the following variables:

```
PORT=3000
GOOGLE_APPLICATION_CREDENTIALS=path/to/your/credentials.json
SMTP_HOST=your-smtp-host
SMTP_PORT=587
SMTP_SECURE=false
SMTP_USER=your-email@example.com
SMTP_PASS=your-password
DEFAULT_FROM_EMAIL=your-email@example.com
```

### Google Sheets API Setup

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project
3. Enable the Google Sheets API
4. Create credentials (Service Account Key)
5. Download the JSON key file
6. Set the path to this file in your `.env` file

### Creating Word Templates

1. Create a Word document (.docx) with your email content
2. Use placeholders like `[Name]`, `[Company]`, etc. that match your Google Sheet column names
3. Save the document

## Usage

1. Start the server:

```bash
npm start
```

2. Open your browser and navigate to `http://localhost:3000`
3. Enter your Google Sheet ID (from the URL of your Google Sheet)
4. Select the sheet containing your contact data
5. Upload your Word template
6. Search for a contact by line number or name
7. Generate an email using the selected template
8. Send the email directly or open in Outlook

## Google Sheet Structure

The application expects your Google Sheet to be structured as follows:

- Multiple sheets, one per region
- Column A contains company names
- Each company name appears once, with multiple rows beneath it for different contacts
- Column E typically contains contact names
- Other columns contain additional contact information

## Troubleshooting

- If you're having trouble connecting to Google Sheets, make sure your credentials are correct and the API is enabled
- If templates aren't working, ensure that the placeholders in your Word document exactly match the column names in your Google Sheet
- For email sending issues, check your SMTP settings in the `.env` file

## License

This project is licensed under the MIT License - see the LICENSE file for details.
