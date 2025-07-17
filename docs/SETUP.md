\# Moralink Automation - Setup Guide



This guide will help you get the Moralink Automation system running for demonstration or development purposes.



\## 📋 Prerequisites



\### Development Environment

\- \*\*Visual Studio 2019\*\* or later (Community Edition is fine)

\- \*\*.NET Framework 4.7.2\*\* or higher

\- \*\*Git\*\* for cloning the repository



\### External Services Required

\- \*\*Moraware JobTracker\*\* account with API access

\- \*\*Google Cloud Platform\*\* account (free tier sufficient for testing)

\- \*\*Google Sheets\*\* for data storage



\## 🚀 Quick Setup (For Demo/Review)



If you just want to see the code structure and UI without connecting to external services:



1\. \*\*Clone the repository\*\*

&nbsp;  ```bash

&nbsp;  git clone \[your-repo-url]

&nbsp;  cd Moralink

&nbsp;  ```



2\. \*\*Open in Visual Studio\*\*

&nbsp;  - Open `MoralinkAutomation.sln`

&nbsp;  - Build the solution (Ctrl+Shift+B)

&nbsp;  - The application will start with default settings



3\. \*\*Explore the interface\*\*

&nbsp;  - The app will run but show connection errors (expected without credentials)

&nbsp;  - Navigate through the Settings tab to see configuration options

&nbsp;  - Review the code structure and implementation patterns



\## 🔧 Full Setup (For Testing/Development)



\### Step 1: Google Cloud Setup



1\. \*\*Create a Google Cloud Project\*\*

&nbsp;  - Go to \[Google Cloud Console](https://console.cloud.google.com/)

&nbsp;  - Create a new project or select existing one



2\. \*\*Enable Google Sheets API\*\*

&nbsp;  - Navigate to "APIs \& Services" > "Library"

&nbsp;  - Search for "Google Sheets API" and enable it



3\. \*\*Create Service Account\*\*

&nbsp;  - Go to "APIs \& Services" > "Credentials"

&nbsp;  - Click "Create Credentials" > "Service Account"

&nbsp;  - Download the JSON key file

&nbsp;  - Rename it to `google-credentials.json`

&nbsp;  - Place it in the project root directory



4\. \*\*Create Google Sheets\*\*

&nbsp;  - Create a new Google Sheet for your schedule

&nbsp;  - Share it with your service account email (found in the JSON file)

&nbsp;  - Note the Sheet ID from the URL



\### Step 2: Moraware Setup



1\. \*\*Get API Access\*\*

&nbsp;  - Ensure your Moraware user has "Execute API Requests" permission

&nbsp;  - Note your database name (e.g., "yourcompany" from yourcompany.moraware.net)



\### Step 3: Application Configuration



1\. \*\*Update Configuration Files\*\*

&nbsp;  - In `Program.cs`, update the default settings in the `AppSettings` constructor:

&nbsp;    ```csharp

&nbsp;    MorawareDatabase = "yourcompany"; // Your Moraware database name

&nbsp;    GoogleSheetId = "your-sheet-id"; // Your Google Sheet ID

&nbsp;    GoogleCredentialsPath = "google-credentials.json";

&nbsp;    ```



2\. \*\*Install Dependencies\*\*

&nbsp;  - Visual Studio should automatically restore NuGet packages

&nbsp;  - If not, run: `Update-Package -Reinstall`



3\. \*\*Build and Run\*\*

&nbsp;  - Build the solution (F6)

&nbsp;  - Run the application (F5)

&nbsp;  - Configure settings through the UI



\### Step 4: Google Apps Script Setup



1\. \*\*Open your Google Sheet\*\*

2\. \*\*Go to Extensions > Apps Script\*\*

3\. \*\*Replace the default code\*\* with contents from `scripts/GoogleAppScript.js`

4\. \*\*Update the configuration\*\*:

&nbsp;  ```javascript

&nbsp;  const MORAWARE\_BASE\_URL = 'https://yourcompany.moraware.net';

&nbsp;  ```

5\. \*\*Save and authorize\*\* the script



\## 🎯 Testing the Integration



\### Basic Functionality Test

1\. \*\*Connection Testing\*\*

&nbsp;  - Open the application

&nbsp;  - Go to Settings

&nbsp;  - Test both Moraware and Google Sheets connections



2\. \*\*Manual Job Scan\*\*

&nbsp;  - Configure a small job ID range (e.g., 100 jobs)

&nbsp;  - Run a manual scan

&nbsp;  - Verify data appears in Google Sheets



3\. \*\*Schedule Automation\*\*

&nbsp;  - Add some test data to your schedule sheet

&nbsp;  - Try the hyperlink generation

&nbsp;  - Test the workflow automation features



\## 🔒 Security Notes



\### For Demonstration/Portfolio

\- Use test accounts and sample data only

\- Never commit real credentials to version control

\- The `.gitignore` file excludes credential files



\### For Production Use

\- Implement proper credential management

\- Use environment variables for sensitive data

\- Regular security audits of API access

\- Monitor usage logs



\## 🛠️ Development Notes



\### Code Structure

```

src/

├── Program.cs              # Main application entry point

├── Models/                 # Data models and DTOs

├── Services/               # API integration services

└── UI/                    # User interface components



scripts/

└── GoogleAppScript.js     # Google Sheets automation



docs/

└── API-Reference.md       # Moraware API documentation

```



\### Key Design Patterns

\- \*\*Repository Pattern\*\*: For data access abstraction

\- \*\*Observer Pattern\*\*: For real-time UI updates

\- \*\*Factory Pattern\*\*: For service initialization

\- \*\*Async/Await\*\*: For non-blocking operations



\### Extensibility Points

\- \*\*New Workflow Stages\*\*: Easily add new production stages

\- \*\*Additional APIs\*\*: Framework supports multiple external services

\- \*\*Custom Reports\*\*: Modular reporting system

\- \*\*Different Sheet Layouts\*\*: Configurable column mappings



\## 📊 Sample Data Structure



\### JobIndex Sheet Format

| Column A | Column B | Column C | Column D | Column E |

|----------|----------|----------|----------|----------|

| Job ID   | Shop Number | Job Name | Digitize Date | Install Dates |



\### AUTOFAB Sheet Format

| Saw Jobs | Saw Desc | Saw Notes | CNC Jobs | CNC Desc | CNC Notes | Polish Jobs | Polish Desc | Polish Notes |

|----------|----------|-----------|----------|----------|-----------|-------------|-------------|-------------|



\## 🐛 Troubleshooting



\### Common Issues



1\. \*\*"Google Sheets API not enabled"\*\*

&nbsp;  - Verify API is enabled in Google Cloud Console

&nbsp;  - Check service account permissions



2\. \*\*"Moraware connection failed"\*\*

&nbsp;  - Verify username/password

&nbsp;  - Check API permissions in Moraware

&nbsp;  - Confirm database name is correct



3\. \*\*"File not found" errors\*\*

&nbsp;  - Ensure credential files are in correct location

&nbsp;  - Check file permissions



4\. \*\*Build errors\*\*

&nbsp;  - Restore NuGet packages

&nbsp;  - Verify .NET Framework version

&nbsp;  - Check all references are resolved



\### Debug Mode

\- Set breakpoints in Visual Studio

\- Use the output console for debugging information

\- Check Windows Event Viewer for application errors



\## 📈 Performance Optimization



\### For Large Datasets

\- Implement batch processing for job scans

\- Use pagination for Google Sheets operations

\- Add caching for frequently accessed data

\- Consider database alternatives for high-volume scenarios



\### Monitoring

\- Built-in logging system

\- Performance counters

\- Error tracking and reporting

\- Usage analytics



\## 🤝 Contributing



This project demonstrates enterprise-level development practices:

\- Clean architecture principles

\- Comprehensive error handling

\- Security best practices

\- Scalable design patterns

\- Professional UI/UX design



For questions or discussions about the implementation, please review the code comments and documentation.



---



\*This setup guide is designed to help technical reviewers and potential employers understand both the project's capabilities and the developer's attention to detail in documentation and user experience.\*

