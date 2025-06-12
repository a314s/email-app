# Email Automation App

A comprehensive email automation application that integrates with Excel files for contact management, email templating, and follow-up tracking.

## Features

- 📊 Excel file integration for contact management
- 📧 Email template processing with Word documents
- 🤖 AI-powered email content improvement using ChatGPT
- 📅 Calendar view for tracking sent emails and follow-ups
- 🔄 Automated follow-up scheduling
- 🏢 Company-specific filtering and views
- 📱 Responsive web interface

## Setup Instructions

### 1. Clone the Repository
```bash
git clone <repository-url>
cd email-automation-app
```

### 2. Install Dependencies
```bash
npm install
```

### 3. Environment Configuration

Copy the example environment file:
```bash
copy .env.example .env
```

Edit the `.env` file and add your configuration:

```env
# Server configuration
PORT=3000

# Google Sheets API
# Replace with your actual credentials path
GOOGLE_APPLICATION_CREDENTIALS=credentials.json

# Email configuration (for direct sending)
SMTP_HOST=smtp.example.com
SMTP_PORT=587
SMTP_SECURE=false
SMTP_USER=your-email@example.com
SMTP_PASS=your-password
DEFAULT_FROM_EMAIL=your-email@example.com

# Anthropic API Key (for AI email improvement)
ANTHROPIC_API_KEY=your-anthropic-api-key-here
```

#### Getting Your API Keys:

**Anthropic API Key:**
1. Go to [Anthropic Console](https://console.anthropic.com/)
2. Sign in or create an account
3. Navigate to API Keys section
4. Click "Create Key"
5. Copy the key and add it to your `.env` file

**Email Configuration:**
- For Gmail, use an App Password instead of your regular password
- Go to Google Account settings → Security → 2-Step Verification → App passwords
- Generate an app password for "Mail"

### 4. Run the Application

Development mode (with auto-restart):
```bash
npm run dev
```

Production mode:
```bash
npm start
```

The application will be available at `http://localhost:3000`

## Usage

1. **Upload Excel Files**: Use the Excel File Configuration section to upload or select existing Excel files
2. **Select Sheet**: Choose the appropriate sheet from your Excel file
3. **Load Contacts**: Click "Load Contacts" to import the data
4. **Search Contacts**: Use line number or name search to find specific contacts
5. **Upload Templates**: Add Word document templates for email generation
6. **Generate Emails**: Select a contact and template to generate personalized emails
7. **AI Improvement**: Use the "Improve Email" feature to enhance your email content with ChatGPT
8. **Calendar View**: Track sent emails and manage follow-ups in the calendar interface

## File Structure

```
email-automation-app/
├── public/           # Frontend files
├── excel/           # Uploaded Excel files
├── templates/       # Email templates (Word documents)
├── data/           # SQLite database
├── server.js       # Main server file
├── database.js     # Database operations
└── package.json    # Dependencies and scripts
```

## Environment Variables

| Variable | Description | Required |
|----------|-------------|----------|
| `ANTHROPIC_API_KEY` | Anthropic API key for AI email improvement | Yes (for AI features) |
| `SMTP_USER` | Email address for sending emails | Yes (for email sending) |
| `SMTP_PASS` | Email password/app password | Yes (for email sending) |
| `SMTP_HOST` | SMTP server hostname | Yes (for email sending) |
| `SMTP_PORT` | SMTP server port | Yes (for email sending) |
| `DEFAULT_FROM_EMAIL` | Default sender email address | Yes (for email sending) |
| `GOOGLE_APPLICATION_CREDENTIALS` | Path to Google Sheets API credentials | Yes (for Google Sheets integration) |
| `PORT` | Server port (default: 3000) | No |

## Security Notes

- Never commit your `.env` file to version control
- Use app passwords for email accounts when possible
- Keep your Anthropic API key secure and monitor usage
- The `.env` file is already included in `.gitignore`
- Always use `.env.example` as a template for new environments
- Regularly rotate API keys and passwords
- Monitor your API usage and set up billing alerts

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is licensed under the MIT License.
