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
# Email Configuration (for sending emails)
EMAIL_USER=your-email@gmail.com
EMAIL_PASSWORD=your-app-password

# OpenAI API Configuration (for AI email improvement)
OPENAI_API_KEY=your-openai-api-key-here

# Server Configuration
PORT=3000
```

#### Getting Your API Keys:

**OpenAI API Key:**
1. Go to [OpenAI API Keys](https://platform.openai.com/api-keys)
2. Sign in or create an account
3. Click "Create new secret key"
4. Copy the key and add it to your `.env` file

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
| `OPENAI_API_KEY` | OpenAI API key for email improvement | Yes (for AI features) |
| `EMAIL_USER` | Email address for sending emails | Yes (for email sending) |
| `EMAIL_PASSWORD` | Email password/app password | Yes (for email sending) |
| `PORT` | Server port (default: 3000) | No |

## Security Notes

- Never commit your `.env` file to version control
- Use app passwords for email accounts when possible
- Keep your OpenAI API key secure and monitor usage
- The `.env` file is already included in `.gitignore`

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is licensed under the MIT License.
