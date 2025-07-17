# Moralink Automation Suite

A comprehensive automation solution that integrates Moraware JobTracker with Google Sheets to streamline workflow management and job tracking for fabrication shops.

## 🌟 Features

### Core Functionality
- **Automated Job Scanning**: Continuously scans Moraware JobTracker for new jobs and activities
- **Google Sheets Integration**: Automatically updates job index and schedule information
- **Hyperlink Generation**: Creates direct links from job numbers to Moraware job pages
- **Schedule Automation**: Moves jobs between workflow stages (Saw → CNC → Polish) based on completion status
- **Daily Reporting**: Generates automated daily activity summaries

### User Interface
- **Modern WinForms Application**: Clean, tabbed interface with real-time status updates
- **Settings Management**: Encrypted settings storage with connection testing
- **Background Processing**: Non-blocking operations with progress tracking
- **Error Handling**: Comprehensive error logging and recovery

### Automation Features
- **Smart Job Detection**: Recognizes completion markers and automatically advances jobs
- **Weekend Handling**: Intelligent scheduling that accounts for business hours
- **Duplicate Prevention**: Prevents duplicate entries across workflow stages
- **Commercial Job Tracking**: Special handling for commercial projects

## 🛠️ Technology Stack

- **Frontend**: C# WinForms (.NET Framework)
- **APIs**: 
  - Moraware JobTracker API v5
  - Google Sheets API v4
- **Authentication**: Google OAuth2 with service accounts
- **Data Storage**: Encrypted XML settings, Google Sheets as database
- **Automation**: Google Apps Script for sheet-side automation

## 📋 Prerequisites

1. **.NET Framework 4.7.2** or higher
2. **Moraware JobTracker** account with API access
3. **Google Cloud Project** with Sheets API enabled
4. **Google Service Account** credentials (JSON file)

## 🚀 Setup Instructions

### 1. Moraware Configuration
1. Ensure your Moraware user has "Execute API Requests" permission
2. Note your Moraware database name (e.g., "yourcompany" from yourcompany.moraware.net)

### 2. Google Sheets Setup
1. Create a Google Cloud Project
2. Enable the Google Sheets API
3. Create a Service Account and download the JSON credentials
4. Share your Google Sheets with the service account email
5. Set up your sheets with the following structure:
   - **JobIndex Sheet**: Columns for Job ID, Shop Number, Job Name, Digitize Date, Install Dates
   - **AUTOFAB Sheet**: Schedule with Saw/CNC/Polish columns

### 3. Application Configuration
1. Clone this repository
2. Build the solution in Visual Studio
3. Run the application and configure settings:
   - Moraware credentials and database name
   - Google Sheets credentials path and sheet IDs
   - Automation preferences

### 4. Google Apps Script Deployment
1. Open your Google Sheet
2. Go to Extensions → Apps Script
3. Paste the contents of `GoogleAppScript.txt`
4. Update the Moraware URL template with your company's subdomain
5. Save and authorize the script

## 📁 Project Structure

```
MoralinkAutomation/
├── src/
│   ├── Program.cs              # Main application and UI
│   ├── Models/                 # Data models and settings
│   └── Services/               # API integration services
├── scripts/
│   └── GoogleAppScript.txt     # Google Sheets automation script
├── docs/
│   └── MorawareAPI.md         # API documentation reference
└── README.md
```

## 🔧 Key Components

### MainForm.cs
- Primary user interface
- Job scanning orchestration
- Settings management
- Real-time logging and status updates

### Settings Management
- Encrypted credential storage
- Connection testing utilities
- Configurable automation parameters

### Job Index System
- Automatic job discovery and cataloging
- Digitize and install date tracking
- Commercial job identification
- Duplicate detection and merging

### Google Sheets Integration
- Batch data updates
- Hyperlink generation
- Schedule synchronization
- Daily activity reporting

## 📊 Workflow Automation

The system automates the following workflow:

1. **Job Discovery**: Scans Moraware for new jobs within specified ranges
2. **Data Extraction**: Pulls job details, activities, and dates
3. **Sheet Updates**: Updates Google Sheets with current information
4. **Link Generation**: Creates hyperlinks from job numbers to Moraware pages
5. **Schedule Management**: Moves completed jobs between workflow stages
6. **Reporting**: Generates daily summaries of completed activities

## ⚙️ Configuration Options

- **Scan Intervals**: Configurable automatic scanning frequency
- **Job Ranges**: Flexible job ID range specification
- **Business Hours**: Weekend and holiday handling
- **Daily Reports**: Automated summary generation and timing
- **Sheet Structure**: Customizable column layouts and ranges

## 🔒 Security Features

- **Encrypted Settings**: All credentials stored with AES-256 encryption
- **Service Account Auth**: Secure Google API access without user intervention
- **Connection Validation**: Built-in testing for all external connections
- **Error Isolation**: Graceful handling of API failures and network issues

## 📈 Performance Benefits

- **Time Savings**: Reduces job lookup time from 15-25 seconds to instant hyperlink clicks
- **Error Reduction**: Eliminates manual data entry and transcription errors
- **Real-time Updates**: Ensures schedule accuracy across all stakeholders
- **Automated Reporting**: Saves hours of daily administrative work

## 🤝 Contributing

This project showcases integration patterns and automation techniques that can be adapted for various ERP and project management systems. The architecture demonstrates:

- Clean separation of concerns
- Robust error handling
- Scalable automation patterns
- User-friendly configuration management

## 📄 License

MIT License - see LICENSE file for details

## 🏭 Use Cases

Perfect for:
- Fabrication shops using Moraware JobTracker
- Teams needing real-time schedule coordination
- Workflows requiring integration between multiple systems
- Organizations seeking to reduce manual data entry

## 📞 Support

For questions about implementation or adaptation to other systems, please open an issue or review the documentation in the `/docs` folder.

---

*This project demonstrates enterprise-level automation capabilities and clean integration patterns suitable for manufacturing and project management environments.*