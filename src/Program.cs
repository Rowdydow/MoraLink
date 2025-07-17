using System;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using System.Threading.Tasks;
using System.IO;
using System.Text;
using System.Security.Cryptography;
using System.Xml.Serialization;
using System.Xml;
using System.Timers;
using Moraware.JobTrackerAPI5;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
// Use aliases to avoid namespace conflicts
using GoogleColor = Google.Apis.Sheets.v4.Data.Color;
using GooglePadding = Google.Apis.Sheets.v4.Data.Padding;
using GoogleData = Google.Apis.Sheets.v4.Data;
using Timer = System.Timers.Timer;

namespace MorawareAutomation
{
    #region Data Models

    [Serializable]
    public class AppSettings
    {
        public string MorawareUsername { get; set; }
        public string MorawarePassword { get; set; }
        public string MorawareDatabase { get; set; }
        public string GoogleSheetScheduleName { get; set; }
        public string GoogleSheetScheduleUrl { get; set; }
        public string GoogleSheetJobIndexName { get; set; }
        public string GoogleSheetJobIndexUrl { get; set; }
        public string GoogleCredentialsPath { get; set; }
        public string GoogleSheetId { get; set; }
        public string JobIndexSheetRange { get; set; }
        public string AutofabSheetName { get; set; }
        public int LastScannedJobNumber { get; set; }
        public bool EnableAutomaticScanning { get; set; }
        public int AutoScanIntervalMinutes { get; set; }
        public DateTime LastScanDate { get; set; }
        public bool DailySummaryEnabled { get; set; }
        public string DailySummaryTime { get; set; } // Format: "HH:MM"

        public AppSettings()
        {
            // Default values
            MorawareUsername = "";
            MorawarePassword = "";
            MorawareDatabase = "yourcompany";
            GoogleSheetScheduleName = "Schedule";
            GoogleSheetScheduleUrl = "";
            GoogleSheetJobIndexName = "JobIndex";
            GoogleSheetJobIndexUrl = "";
            GoogleCredentialsPath = "example-credential.json";
            GoogleSheetId = "Your_GoogleSheetId_Here";
            JobIndexSheetRange = "JobIndex!B3:G";
            AutofabSheetName = "AUTOFAB";
            LastScannedJobNumber = 19000;
            EnableAutomaticScanning = false;
            AutoScanIntervalMinutes = 120; // Default: scan every 2 hours
            LastScanDate = DateTime.MinValue;
            DailySummaryEnabled = false;
            DailySummaryTime = "17:00"; // Default: 5:00 PM
        }
    }

    public class JobIndexEntry
    {
        public string JobNumber { get; set; }
        public string ShopNumber { get; set; }
        public string JobName { get; set; }
        public DateTime? LatestDigitizeDate { get; set; }
        public List<DateTime> InstallDates { get; set; } = new List<DateTime>();
        public string FormattedInstallDates { get; set; } = "";

        public override bool Equals(object obj)
        {
            if (obj is JobIndexEntry other)
            {
                return JobNumber == other.JobNumber;
            }
            return false;
        }

        public override int GetHashCode()
        {
            int hash = 17;
            hash = hash * 23 + (JobNumber ?? "").GetHashCode();
            return hash;
        }
    }

    public class DailyActivitySummary
    {
        public DateTime Date { get; set; }
        public List<string> CompletedSawJobs { get; set; } = new List<string>();
        public List<string> CompletedCncJobs { get; set; } = new List<string>();
        public List<string> CompletedPolishJobs { get; set; } = new List<string>();
    }

    public class DateSection
    {
        public int StartRow { get; set; }
        public int EndRow { get; set; }
        public string DateText { get; set; }
        public DateTime Date { get; set; }
    }

    #endregion

    #region Settings Management

    public static class SettingsManager
    {
        private static readonly string SettingsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MorawareAutomation",
            "settings.xml");

        private static readonly byte[] EncryptionKey = Encoding.UTF8.GetBytes("MorawareAutomation2025SecureKey!"); // 32 bytes for AES-256

        public static AppSettings LoadSettings()
        {
            try
            {
                // Create directory if it doesn't exist
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsFilePath));

                if (!File.Exists(SettingsFilePath))
                {
                    return new AppSettings();
                }

                string encryptedData = File.ReadAllText(SettingsFilePath);
                string decryptedXml = DecryptString(encryptedData);

                XmlSerializer serializer = new XmlSerializer(typeof(AppSettings));
                using (StringReader reader = new StringReader(decryptedXml))
                {
                    return (AppSettings)serializer.Deserialize(reader);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading settings: {ex.Message}\nDefault settings will be used.", 
                    "Settings Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return new AppSettings();
            }
        }

        public static void SaveSettings(AppSettings settings)
        {
            try
            {
                // Create directory if it doesn't exist
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsFilePath));

                XmlSerializer serializer = new XmlSerializer(typeof(AppSettings));
                XmlWriterSettings xmlSettings = new XmlWriterSettings
                {
                    Indent = true,
                    IndentChars = "  ",
                    NewLineChars = "\r\n",
                    NewLineHandling = NewLineHandling.Replace
                };

                StringBuilder sb = new StringBuilder();
                using (XmlWriter writer = XmlWriter.Create(sb, xmlSettings))
                {
                    serializer.Serialize(writer, settings);
                }

                string encryptedData = EncryptString(sb.ToString());
                File.WriteAllText(SettingsFilePath, encryptedData);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving settings: {ex.Message}", 
                    "Settings Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string EncryptString(string plainText)
        {
            byte[] iv = new byte[16]; // 16 bytes for AES
            using (var rng = new RNGCryptoServiceProvider())
            {
                rng.GetBytes(iv);
            }

            using (Aes aes = Aes.Create())
            {
                aes.Key = EncryptionKey;
                aes.IV = iv;

                ICryptoTransform encryptor = aes.CreateEncryptor(aes.Key, aes.IV);

                using (MemoryStream memoryStream = new MemoryStream())
                {
                    // Write the IV first
                    memoryStream.Write(iv, 0, iv.Length);

                    using (CryptoStream cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write))
                    using (StreamWriter streamWriter = new StreamWriter(cryptoStream))
                    {
                        streamWriter.Write(plainText);
                    }
                    return Convert.ToBase64String(memoryStream.ToArray());
                }
            }
        }

        private static string DecryptString(string cipherText)
        {
            byte[] fullCipher = Convert.FromBase64String(cipherText);

            // Extract the IV from the cipher text
            byte[] iv = new byte[16];
            byte[] cipher = new byte[fullCipher.Length - 16];

            Buffer.BlockCopy(fullCipher, 0, iv, 0, iv.Length);
            Buffer.BlockCopy(fullCipher, iv.Length, cipher, 0, cipher.Length);

            using (Aes aes = Aes.Create())
            {
                aes.Key = EncryptionKey;
                aes.IV = iv;

                ICryptoTransform decryptor = aes.CreateDecryptor(aes.Key, aes.IV);

                using (MemoryStream memoryStream = new MemoryStream(cipher))
                using (CryptoStream cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read))
                using (StreamReader streamReader = new StreamReader(cryptoStream))
                {
                    return streamReader.ReadToEnd();
                }
            }
        }
    }

    #endregion

    #region Main Form

    public class MainForm : Form
    {
        private AppSettings settings;
        private Connection morawareConnection;
        private SheetsService sheetsService;
        private HashSet<JobIndexEntry> jobIndex = new HashSet<JobIndexEntry>();

        // UI Controls
        private RichTextBox outputTextBox;
        private TabControl tabControl;
        private Button startScanButton;
        private Button stopScanButton;
        private Button settingsButton;
        private Button viewJobIndexButton;
        private Button dailySummaryButton;
        private Label statusLabel;
        private Timer autoScanTimer;
        private Timer dailySummaryTimer;

        public MainForm()
        {
            settings = SettingsManager.LoadSettings();
            InitializeComponents();
            InitializeGoogleSheetsService();
            SetupTimers();
        }

        private void InitializeComponents()
        {
            this.Size = new Size(900, 650);
            this.Text = "Moralink Automation";
            this.Icon = SystemIcons.Application;

            // Main layout
            tabControl = new TabControl { Dock = DockStyle.Fill };
            
            // Dashboard tab
            TabPage dashboardTab = new TabPage("Dashboard");
            var dashboardPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                RowCount = 3,
                ColumnCount = 2
            };

            // Control panel
            var controlPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle
            };

            var controlLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                RowCount = 5,
                ColumnCount = 1
            };

            startScanButton = new Button
            {
                Text = "Start Manual Scan",
                Dock = DockStyle.Fill,
                Height = 40,
                BackColor = Color.LightGreen,
                FlatStyle = FlatStyle.Flat
            };
            startScanButton.Click += StartScanButton_Click;

            stopScanButton = new Button
            {
                Text = "Stop Running Scan",
                Dock = DockStyle.Fill,
                Height = 40,
                Enabled = false,
                BackColor = Color.LightCoral,
                FlatStyle = FlatStyle.Flat
            };
            stopScanButton.Click += StopScanButton_Click;

            settingsButton = new Button
            {
                Text = "Settings",
                Dock = DockStyle.Fill,
                Height = 40,
                BackColor = Color.LightBlue,
                FlatStyle = FlatStyle.Flat
            };
            settingsButton.Click += SettingsButton_Click;

            viewJobIndexButton = new Button
            {
                Text = "View Job Index",
                Dock = DockStyle.Fill,
                Height = 40,
                BackColor = Color.LightYellow,
                FlatStyle = FlatStyle.Flat
            };
            viewJobIndexButton.Click += ViewJobIndexButton_Click;

            dailySummaryButton = new Button
            {
                Text = "Generate Daily Summary",
                Dock = DockStyle.Fill,
                Height = 40,
                BackColor = Color.LightGray,
                FlatStyle = FlatStyle.Flat
            };
            dailySummaryButton.Click += DailySummaryButton_Click;

            controlLayout.Controls.Add(startScanButton, 0, 0);
            controlLayout.Controls.Add(stopScanButton, 0, 1);
            controlLayout.Controls.Add(settingsButton, 0, 2);
            controlLayout.Controls.Add(viewJobIndexButton, 0, 3);
            controlLayout.Controls.Add(dailySummaryButton, 0, 4);

            controlLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 20F));
            controlLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 20F));
            controlLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 20F));
            controlLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 20F));
            controlLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 20F));

            controlPanel.Controls.Add(controlLayout);

            // Status panel
            var statusPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle
            };

            var statusLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                RowCount = 1,
                ColumnCount = 1
            };

            statusLabel = new Label
            {
                Text = GetStatusText(),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9F),
                AutoSize = false
            };

            statusLayout.Controls.Add(statusLabel, 0, 0);
            statusPanel.Controls.Add(statusLayout);

            // Output panel
            outputTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 10F),
                BackColor = Color.White
            };

            // Add controls to dashboard
            dashboardPanel.Controls.Add(controlPanel, 0, 0);
            dashboardPanel.Controls.Add(statusPanel, 1, 0);
            dashboardPanel.Controls.Add(outputTextBox, 0, 1);
            dashboardPanel.SetColumnSpan(outputTextBox, 2);
            dashboardPanel.SetRowSpan(controlPanel, 1);
            dashboardPanel.SetRowSpan(statusPanel, 1);

            dashboardPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));
            dashboardPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
            dashboardPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 30F));
            dashboardPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 70F));

            dashboardTab.Controls.Add(dashboardPanel);
            
            // Add tabs to tab control
            tabControl.Controls.Add(dashboardTab);
            
            this.Controls.Add(tabControl);
        }

        private void SetupTimers()
        {
            // Auto-scan timer
            autoScanTimer = new Timer(settings.AutoScanIntervalMinutes * 60 * 1000); // Convert minutes to milliseconds
            autoScanTimer.Elapsed += AutoScanTimer_Elapsed;
            autoScanTimer.AutoReset = true;
            
            if (settings.EnableAutomaticScanning)
            {
                autoScanTimer.Start();
                LogOutput("Automatic scanning enabled.");
            }
            
            // Daily summary timer
            dailySummaryTimer = new Timer();
            SetupDailySummaryTimer();
            
            if (settings.DailySummaryEnabled)
            {
                dailySummaryTimer.Start();
                LogOutput("Daily summary report enabled.");
            }
        }

        private void SetupDailySummaryTimer()
        {
            dailySummaryTimer.Stop();
            
            try
            {
                // Parse time from settings (format: "HH:MM")
                string[] timeParts = settings.DailySummaryTime.Split(':');
                int hour = int.Parse(timeParts[0]);
                int minute = int.Parse(timeParts[1]);
                
                // Calculate time until next run
                DateTime now = DateTime.Now;
                DateTime scheduledTime = new DateTime(now.Year, now.Month, now.Day, hour, minute, 0);
                
                // If scheduled time has already passed today, schedule for tomorrow
                if (scheduledTime <= now)
                {
                    scheduledTime = scheduledTime.AddDays(1);
                }
                
                double intervalMs = (scheduledTime - now).TotalMilliseconds;
                
                dailySummaryTimer.Interval = intervalMs;
                dailySummaryTimer.Elapsed += DailySummaryTimer_Elapsed;
                dailySummaryTimer.AutoReset = false; // Only run once, then reschedule
                
                LogOutput($"Daily summary scheduled for {scheduledTime.ToString("MM/dd/yyyy HH:mm")}");
            }
            catch (Exception ex)
            {
                LogOutput($"Error setting up daily summary timer: {ex.Message}");
            }
        }

        private void DailySummaryTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                LogOutput("Generating daily activity summary...");
                // Use Task.Run to handle async operations in a non-async event handler
                Task.Run(async () => 
                {
                    await GenerateDailyActivitySummary();
                    
                    // Reschedule for next day
                    SetupDailySummaryTimer();
                    dailySummaryTimer.Start();
                }).Wait();
            }
            catch (Exception ex)
            {
                LogOutput($"Error during daily summary: {ex.Message}");
            }
        }

        private void AutoScanTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                LogOutput("Starting automated job scan...");
                // Use Task.Run to handle async operations in a non-async event handler
                Task.Run(async () => 
                {
                    await PerformJobScan(true);
                    UpdateStatusLabel();
                }).Wait();
            }
            catch (Exception ex)
            {
                LogOutput($"Error during automated scan: {ex.Message}");
            }
        }

        private string GetStatusText()
        {
            StringBuilder status = new StringBuilder();
            
            status.AppendLine("Moralink Automation Status");
            status.AppendLine("------------------------");
            status.AppendLine($"Connected to: {settings.MorawareDatabase}.moraware.net");
            status.AppendLine($"Last scanned job: {settings.LastScannedJobNumber}");
            status.AppendLine($"Last scan date: {(settings.LastScanDate == DateTime.MinValue ? "Never" : settings.LastScanDate.ToString("MM/dd/yyyy HH:mm:ss"))}");
            status.AppendLine($"Auto scanning: {(settings.EnableAutomaticScanning ? "Enabled" : "Disabled")}");
            
            if (settings.EnableAutomaticScanning)
            {
                status.AppendLine($"Scan interval: Every {settings.AutoScanIntervalMinutes} minutes");
            }
            
            status.AppendLine($"Daily summary: {(settings.DailySummaryEnabled ? $"Enabled (at {settings.DailySummaryTime})" : "Disabled")}");
            
            return status.ToString();
        }

        private void UpdateStatusLabel()
        {
            if (statusLabel.InvokeRequired)
            {
                statusLabel.Invoke(new Action(UpdateStatusLabel));
                return;
            }
            
            statusLabel.Text = GetStatusText();
        }

        private void InitializeGoogleSheetsService()
        {
            try
            {
                GoogleCredential credential;
                using (var stream = new FileStream(settings.GoogleCredentialsPath,
                    FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream)
                        .CreateScoped(new[] { SheetsService.Scope.Spreadsheets });
                }

                sheetsService = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Moralink Automation"
                });

                LogOutput("Google Sheets service initialized successfully.");
            }
            catch (Exception ex)
            {
                LogOutput($"Error initializing Google Sheets: {ex.Message}");
            }
        }

        private async Task<List<JobIndexEntry>> GetExistingEntries()
        {
            try
            {
                var existingEntries = new List<JobIndexEntry>();
                var request = sheetsService.Spreadsheets.Values.Get(settings.GoogleSheetId, settings.JobIndexSheetRange);
                var response = await request.ExecuteAsync();

                if (response.Values != null)
                {
                    foreach (var row in response.Values)
                    {
                        var entry = new JobIndexEntry();

                        // Handle JobNumber
                        entry.JobNumber = row.Count > 0 && row[0] != null ? row[0].ToString() : null;

                        // Handle ShopNumber
                        entry.ShopNumber = row.Count > 1 && row[1] != null ? row[1].ToString() : null;

                        // Handle JobName
                        entry.JobName = row.Count > 2 && row[2] != null ? row[2].ToString() : null;

                        // Handle LatestDigitizeDate
                        if (row.Count > 3 && row[3] != null && !string.IsNullOrEmpty(row[3].ToString()))
                        {
                            DateTime digitizeDate;
                            if (DateTime.TryParse(row[3].ToString(), out digitizeDate))
                            {
                                entry.LatestDigitizeDate = digitizeDate;
                            }
                        }

                        // Handle Install Dates (as a comma-separated string)
                        if (row.Count > 4 && row[4] != null && !string.IsNullOrEmpty(row[4].ToString()))
                        {
                            entry.FormattedInstallDates = row[4].ToString();
                            
                            // Parse individual dates
                            string[] dateStrings = entry.FormattedInstallDates.Split(',');
                            foreach (var dateStr in dateStrings)
                            {
                                DateTime installDate;
                                if (DateTime.TryParse(dateStr.Trim(), out installDate))
                                {
                                    entry.InstallDates.Add(installDate);
                                }
                            }
                        }

                        // Update the highest job number we've seen
                        if (!string.IsNullOrEmpty(entry.JobNumber) && int.TryParse(entry.JobNumber, out int jobNumber))
                        {
                            settings.LastScannedJobNumber = Math.Max(settings.LastScannedJobNumber, jobNumber);
                        }

                        existingEntries.Add(entry);
                    }
                }

                return existingEntries;
            }
            catch (Exception ex)
            {
                LogOutput($"Error reading existing entries: {ex.Message}");
                return new List<JobIndexEntry>();
            }
        }

        private async Task<List<string>> GetCommercialJobIds()
        {
            try
            {
                var commercialJobIds = new List<string>();
                var existingEntries = await GetExistingEntries();
                
                foreach (var entry in existingEntries)
                {
                    if (entry.JobName != null && entry.JobName.StartsWith("Commercial", StringComparison.OrdinalIgnoreCase))
                    {
                        commercialJobIds.Add(entry.JobNumber);
                    }
                }
                
                return commercialJobIds;
            }
            catch (Exception ex)
            {
                LogOutput($"Error getting commercial job IDs: {ex.Message}");
                return new List<string>();
            }
        }

        private async Task UpdateJobIndex()
        {
            try
            {
                var existingEntries = await GetExistingEntries();
                LogOutput($"Found {existingEntries.Count} existing entries in Job Index.");

                var mergedEntries = new HashSet<JobIndexEntry>();

                // First, add all existing entries
                foreach (var existing in existingEntries)
                {
                    mergedEntries.Add(existing);
                }

                // Then update or add new entries
                foreach (var newEntry in jobIndex)
                {
                    var existingEntry = mergedEntries.FirstOrDefault(e => e.Equals(newEntry));
                    if (existingEntry != null)
                    {
                        // Update digitize date if new entry has more recent date
                        if (newEntry.LatestDigitizeDate.HasValue)
                        {
                            if (!existingEntry.LatestDigitizeDate.HasValue ||
                                newEntry.LatestDigitizeDate.Value > existingEntry.LatestDigitizeDate.Value)
                            {
                                existingEntry.LatestDigitizeDate = newEntry.LatestDigitizeDate;
                            }
                        }

                        // Update or add install dates
                        var today = DateTime.Today;
                        var combinedDates = new HashSet<DateTime>(existingEntry.InstallDates);
                        
                        // Add only future and today's install dates
                        foreach (var date in newEntry.InstallDates)
                        {
                            if (date >= today)
                            {
                                combinedDates.Add(date);
                            }
                        }
                        
                        existingEntry.InstallDates = combinedDates.OrderBy(d => d).ToList();
                        
                        // Format the dates as a comma-separated string
                        if (existingEntry.InstallDates.Any())
                        {
                            existingEntry.FormattedInstallDates = string.Join(", ", 
                                existingEntry.InstallDates.Select(d => d.ToString("MM/dd/yyyy")));
                        }
                        else
                        {
                            existingEntry.FormattedInstallDates = "";
                        }
                    }
                    else
                    {
                        // Filter to only include future and today's install dates
                        var today = DateTime.Today;
                        newEntry.InstallDates = newEntry.InstallDates
                            .Where(d => d >= today)
                            .OrderBy(d => d)
                            .ToList();
                            
                        // Format the dates
                        if (newEntry.InstallDates.Any())
                        {
                            newEntry.FormattedInstallDates = string.Join(", ", 
                                newEntry.InstallDates.Select(d => d.ToString("MM/dd/yyyy")));
                        }
                        
                        mergedEntries.Add(newEntry);
                    }
                }

                LogOutput($"Preparing to update sheet with {mergedEntries.Count} entries.");

                // Prepare the data for the sheet
                var valueRange = new GoogleData.ValueRange();
                var values = new List<IList<object>>();

                // Convert merged entries to sheet values
                foreach (var entry in mergedEntries)
                {
                    var digitizeDateString = entry.LatestDigitizeDate.HasValue ?
                        entry.LatestDigitizeDate.Value.ToString("MM/dd/yyyy") : "";

                    values.Add(new List<object>
                    {
                        entry.JobNumber,
                        entry.ShopNumber,
                        entry.JobName,
                        digitizeDateString,
                        entry.FormattedInstallDates
                    });
                }

                valueRange.Values = values;

                // Clear existing data
                var clearRequest = sheetsService.Spreadsheets.Values.Clear(
                    new GoogleData.ClearValuesRequest(), settings.GoogleSheetId, settings.JobIndexSheetRange);
                await clearRequest.ExecuteAsync();

                // Write new data
                var updateRequest = sheetsService.Spreadsheets.Values.Update(valueRange, settings.GoogleSheetId, settings.JobIndexSheetRange);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                await updateRequest.ExecuteAsync();

                LogOutput($"Job Index updated successfully. Total entries: {mergedEntries.Count}");
            }
            catch (Exception ex)
            {
                LogOutput($"Error updating Google Sheet: {ex.Message}");
            }
        }

        private async Task<List<DateSection>> FindDateSections()
        {
            try
            {
                LogOutput("Searching for date sections in AUTOFAB sheet...");
                var range = $"{settings.AutofabSheetName}!A1:Z1000";
                var request = sheetsService.Spreadsheets.Values.Get(settings.GoogleSheetId, range);
                var response = await request.ExecuteAsync();
                var dateSections = new List<DateSection>();

                if (response.Values == null)
                {
                    LogOutput("No values found in AUTOFAB sheet.");
                    return dateSections;
                }

                int currentRow = 1;
                foreach (var row in response.Values)
                {
                    if (row.Count > 0)
                    {
                        var cellValue = row[0].ToString();
                        
                        // Check if this is a date row by looking for common patterns
                        // Example: Monday - 03/07/25
                        if (!string.IsNullOrEmpty(cellValue) && 
                            (cellValue.Contains("/") || 
                             cellValue.Contains("Monday") || 
                             cellValue.Contains("Tuesday") || 
                             cellValue.Contains("Wednesday") || 
                             cellValue.Contains("Thursday") || 
                             cellValue.Contains("Friday")))
                        {
                            try
                            {
                                // Try to extract a date from the string
                                DateTime rowDate = DateTime.MinValue;
                                string dateText = "";
                                
                                // Common format: "Day - MM/DD/YY"
                                var dateMatch = System.Text.RegularExpressions.Regex.Match(cellValue, @"\d{1,2}/\d{1,2}/\d{2,4}");
                                if (dateMatch.Success)
                                {
                                    dateText = dateMatch.Value;
                                    rowDate = DateTime.Parse(dateText);
                                }
                                
                                if (rowDate != DateTime.MinValue)
                                {
                                    LogOutput($"Found date row at {currentRow}: {cellValue}");
                                    
                                    var section = new DateSection
                                    {
                                        StartRow = currentRow + 1,
                                        DateText = cellValue,
                                        Date = rowDate
                                    };

                                    if (dateSections.Any())
                                    {
                                        dateSections.Last().EndRow = currentRow - 1;
                                    }

                                    dateSections.Add(section);
                                }
                            }
                            catch
                            {
                                // Not a valid date, continue
                            }
                        }
                    }
                    currentRow++;
                }

                if (dateSections.Any())
                {
                    dateSections.Last().EndRow = currentRow - 1;
                    LogOutput($"Found {dateSections.Count} date sections");
                    foreach (var section in dateSections)
                    {
                        LogOutput($"Section: {section.DateText} (Rows {section.StartRow}-{section.EndRow})");
                    }
                }
                else
                {
                    LogOutput("No date sections found");
                }

                return dateSections;
            }
            catch (Exception ex)
            {
                LogOutput($"Error finding date sections: {ex.Message}");
                return new List<DateSection>();
            }
        }

        private async Task<Dictionary<string, List<(string JobNumber, string JobName)>>> GetCompletedJobs(DateSection section)
        {
            try
            {
                var completedJobs = new Dictionary<string, List<(string, string)>>
                {
                    { "Saw", new List<(string, string)>() },
                    { "CNC", new List<(string, string)>() },
                    { "Polish", new List<(string, string)>() }
                };

                // 1. Get completed Saw jobs (columns A-C)
                var sawRange = $"{settings.AutofabSheetName}!A{section.StartRow}:C{section.EndRow}";
                var sawData = await GetSheetDataAsync(sawRange);

                foreach (var row in sawData)
                {
                    // Check if there's valid job data
                    if (row.Count >= 2 && row[0] != null && !string.IsNullOrWhiteSpace(row[0].ToString()))
                    {
                        var jobNumber = row[0].ToString();
                        var jobName = row.Count > 1 && row[1] != null ? row[1].ToString() : "";
                        var rowIndex = section.StartRow + sawData.IndexOf(row);
                        
                        // We'd need to check if this job is marked as completed
                        // For now, assume completed based on any value in column C
                        if (row.Count > 2 && row[2] != null && !string.IsNullOrWhiteSpace(row[2].ToString()))
                        {
                            completedJobs["Saw"].Add((jobNumber, jobName));
                        }
                    }
                }

                // 2. Get completed CNC jobs (columns D-F)
                var cncRange = $"{settings.AutofabSheetName}!D{section.StartRow}:F{section.EndRow}";
                var cncData = await GetSheetDataAsync(cncRange);

                foreach (var row in cncData)
                {
                    if (row.Count >= 2 && row[0] != null && !string.IsNullOrWhiteSpace(row[0].ToString()))
                    {
                        var jobNumber = row[0].ToString();
                        var jobName = row.Count > 1 && row[1] != null ? row[1].ToString() : "";
                        var rowIndex = section.StartRow + cncData.IndexOf(row);
                        
                        // Check for completion marker
                        if (row.Count > 2 && row[2] != null && !string.IsNullOrWhiteSpace(row[2].ToString()))
                        {
                            completedJobs["CNC"].Add((jobNumber, jobName));
                        }
                    }
                }

                // 3. Get completed Polish jobs (columns G-I)
                var polishRange = $"{settings.AutofabSheetName}!G{section.StartRow}:I{section.EndRow}";
                var polishData = await GetSheetDataAsync(polishRange);

                foreach (var row in polishData)
                {
                    if (row.Count >= 2 && row[0] != null && !string.IsNullOrWhiteSpace(row[0].ToString()))
                    {
                        var jobNumber = row[0].ToString();
                        var jobName = row.Count > 1 && row[1] != null ? row[1].ToString() : "";
                        var rowIndex = section.StartRow + polishData.IndexOf(row);
                        
                        // Check for completion marker
                        if (row.Count > 2 && row[2] != null && !string.IsNullOrWhiteSpace(row[2].ToString()))
                        {
                            completedJobs["Polish"].Add((jobNumber, jobName));
                        }
                    }
                }

                LogOutput($"Found {completedJobs["Saw"].Count} completed Saw jobs");
                LogOutput($"Found {completedJobs["CNC"].Count} completed CNC jobs");
                LogOutput($"Found {completedJobs["Polish"].Count} completed Polish jobs");

                return completedJobs;
            }
            catch (Exception ex)
            {
                LogOutput($"Error getting completed jobs: {ex.Message}");
                return new Dictionary<string, List<(string, string)>>();
            }
        }

        private async Task GenerateDailyActivitySummary()
        {
            try
            {
                LogOutput("Generating daily activity summary...");
                
                // Find today's date section in the AUTOFAB sheet
                var dateSections = await FindDateSections();
                var today = DateTime.Today;
                
                var todaySection = dateSections.FirstOrDefault(s => s.Date.Date == today);
                if (todaySection == null)
                {
                    LogOutput("No section found for today's date in the AUTOFAB sheet.");
                    return;
                }
                
                // Get completed jobs for today
                var completedJobs = await GetCompletedJobs(todaySection);
                
                // Create and save summary
                var summary = new DailyActivitySummary
                {
                    Date = today,
                    CompletedSawJobs = completedJobs["Saw"].Select(j => $"{j.JobNumber} - {j.JobName}").ToList(),
                    CompletedCncJobs = completedJobs["CNC"].Select(j => $"{j.JobNumber} - {j.JobName}").ToList(),
                    CompletedPolishJobs = completedJobs["Polish"].Select(j => $"{j.JobNumber} - {j.JobName}").ToList()
                };
                
                // Save to file
                string summaryDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "MoralinkReports");
                
                Directory.CreateDirectory(summaryDirectory);
                
                string summaryPath = Path.Combine(
                    summaryDirectory, 
                    $"DailySummary_{today:yyyy-MM-dd}.txt");
                
                using (var writer = new StreamWriter(summaryPath))
                {
                    writer.WriteLine($"Daily Activity Summary for {today:MM/dd/yyyy}");
                    writer.WriteLine("=============================================");
                    writer.WriteLine();
                    
                    writer.WriteLine("COMPLETED SAW JOBS:");
                    writer.WriteLine("-------------------");
                    if (summary.CompletedSawJobs.Any())
                    {
                        foreach (var job in summary.CompletedSawJobs)
                        {
                            writer.WriteLine($"- {job}");
                        }
                    }
                    else
                    {
                        writer.WriteLine("No completed saw jobs today.");
                    }
                    writer.WriteLine();
                    
                    writer.WriteLine("COMPLETED CNC JOBS:");
                    writer.WriteLine("-------------------");
                    if (summary.CompletedCncJobs.Any())
                    {
                        foreach (var job in summary.CompletedCncJobs)
                        {
                            writer.WriteLine($"- {job}");
                        }
                    }
                    else
                    {
                        writer.WriteLine("No completed CNC jobs today.");
                    }
                    writer.WriteLine();
                    
                    writer.WriteLine("COMPLETED POLISH JOBS:");
                    writer.WriteLine("----------------------");
                    if (summary.CompletedPolishJobs.Any())
                    {
                        foreach (var job in summary.CompletedPolishJobs)
                        {
                            writer.WriteLine($"- {job}");
                        }
                    }
                    else
                    {
                        writer.WriteLine("No completed polish jobs today.");
                    }
                }
                
                LogOutput($"Daily summary saved to: {summaryPath}");
                
                // Open the file in default text editor
                System.Diagnostics.Process.Start(summaryPath);
            }
            catch (Exception ex)
            {
                LogOutput($"Error generating daily summary: {ex.Message}");
            }
        }

        private async Task<IList<IList<object>>> GetSheetDataAsync(string range)
        {
            LogOutput($"Getting sheet data for range: {range}");
            var request = sheetsService.Spreadsheets.Values.Get(settings.GoogleSheetId, range);
            var response = await request.ExecuteAsync();
            if (response.Values == null)
            {
                LogOutput("No data found in range");
                return new List<IList<object>>();
            }
            LogOutput($"Retrieved {response.Values.Count} rows of data");
            return response.Values;
        }

        private async Task PerformJobScan(bool isAutomated = false)
        {
            startScanButton.Enabled = false;
            stopScanButton.Enabled = true;
            
            if (!isAutomated)
            {
                outputTextBox.Clear();
            }
            
            jobIndex.Clear();  // Clear previous job index entries

            try
            {
                // Connect to Moraware
                LogOutput("Connecting to Moraware...");
                morawareConnection = new Connection(
                    $"https://{settings.MorawareDatabase}.moraware.net/api.aspx", 
                    settings.MorawareUsername, 
                    settings.MorawarePassword);
                morawareConnection.Connect();
                LogOutput("Connected successfully!\n");

                var jobsToCheck = new HashSet<int>();
                var jobData = new Dictionary<int, (string JobName, string FormName, List<DateTime?> DigitizeDates, List<DateTime?> InstallDates)>();

                // Calculate job range based on last scanned job
                int startId = Math.Max(1, settings.LastScannedJobNumber - 1250);
                int endId = settings.LastScannedJobNumber + 100;
                
                LogOutput($"Adding job range {startId} to {endId} to scan queue...");
                for (int i = startId; i <= endId; i++)
                {
                    jobsToCheck.Add(i);
                }

                // Add Commercial jobs
                LogOutput("Adding Commercial jobs from JobIndex...");
                var commercialJobIds = await GetCommercialJobIds();
                foreach (var jobIdStr in commercialJobIds)
                {
                    if (int.TryParse(jobIdStr, out int jobId))
                    {
                        jobsToCheck.Add(jobId);
                        LogOutput($"Added Commercial job {jobId} to scan queue");
                    }
                }

                // Track the highest job number we encounter
                int highestJobNumber = settings.LastScannedJobNumber;

                // Check all jobs
                LogOutput($"Starting scan of {jobsToCheck.Count} jobs...");
                int processedCount = 0;
                foreach (var jobId in jobsToCheck)
                {
                    try
                    {
                        processedCount++;
                        if (processedCount % 50 == 0)
                        {
                            LogOutput($"Progress: {processedCount}/{jobsToCheck.Count} jobs processed");
                        }
                        
                        var job = morawareConnection.GetJob(jobId);

                        if (job != null)
                        {
                            highestJobNumber = Math.Max(highestJobNumber, jobId);
                            
                            var forms = morawareConnection.GetJobForms(jobId, true);
                            var activities = morawareConnection.GetJobActivities(jobId, true, true);
                            var digitizeActivities = activities.Where(a => a.JobActivityTypeName == "Digitize").ToList();
                            var installActivities = activities.Where(a => a.JobActivityTypeName == "Install").ToList();
                            var formName = forms.FirstOrDefault()?.JobFormName ?? "No Form Name";
                            var jobName = job.JobName ?? "No Job Name";
                            var latestDigitizeDate = (DateTime?)null;

                            // Extract shop number from job name
                            string shopNumber = "";
                            var shopNumberMatch = System.Text.RegularExpressions.Regex
                                .Match(jobName, @"\d{2}-\d{3,4}");

                            if (shopNumberMatch.Success)
                            {
                                shopNumber = shopNumberMatch.Value;
                            }

                            if (digitizeActivities.Any())
                            {
                                latestDigitizeDate = digitizeActivities.Max(a => a.StartDate);
                            }

                            var entry = new JobIndexEntry
                            {
                                JobNumber = jobId.ToString(),
                                ShopNumber = shopNumber,
                                JobName = jobName,
                                LatestDigitizeDate = latestDigitizeDate
                            };

                            // Get all install dates
                            if (installActivities.Any())
                            {
                                foreach (var activity in installActivities)
                                {
                                    entry.InstallDates.Add(activity.StartDate);
                                }
                            }

                            jobIndex.Add(entry);

                            jobData[jobId] = (
                                jobName,
                                formName,
                                digitizeActivities.Select(a => a.StartDate).ToList(),
                                installActivities.Select(a => a.StartDate).ToList()
                            );
                        }
                    }
                    catch (Exception ex)
                    {
                        LogOutput($"Error processing job {jobId}: {ex.Message}");
                    }
                }

                // Update the last scanned job number in settings
                settings.LastScannedJobNumber = highestJobNumber;
                settings.LastScanDate = DateTime.Now;
                SettingsManager.SaveSettings(settings);
                
                // Update Google Sheet with collected job information
                await UpdateJobIndex();
                LogOutput($"Job scan completed. Processed {processedCount} jobs.");
                
                morawareConnection.Disconnect();
                LogOutput("Disconnected from Moraware successfully!");
                
                // Update status display
                UpdateStatusLabel();
            }
            catch (Exception ex)
            {
                LogOutput($"\nCritical Error: {ex.Message}");
                LogOutput($"Stack Trace: {ex.StackTrace}");
            }
            finally
            {
                startScanButton.Enabled = true;
                stopScanButton.Enabled = false;
            }
        }

        private void LogOutput(string message)
        {
            if (outputTextBox.InvokeRequired)
            {
                outputTextBox.Invoke(new Action(() => LogOutput(message)));
                return;
            }

            outputTextBox.AppendText(message + "\n");
            outputTextBox.ScrollToCaret();
        }

        private void StartScanButton_Click(object sender, EventArgs e)
        {
            Task.Run(async () => await PerformJobScan());
        }

        private void StopScanButton_Click(object sender, EventArgs e)
        {
            // This would need to implement cancellation tokens to properly stop a scan
            LogOutput("Stop requested. Will finish current job and then stop.");
        }

        private void SettingsButton_Click(object sender, EventArgs e)
        {
            var settingsForm = new SettingsForm(settings);
            if (settingsForm.ShowDialog() == DialogResult.OK)
            {
                // Update settings
                settings = settingsForm.Settings;
                SettingsManager.SaveSettings(settings);
                
                // Reinitialize with new settings
                InitializeGoogleSheetsService();
                
                // Update timers
                autoScanTimer.Interval = settings.AutoScanIntervalMinutes * 60 * 1000;
                if (settings.EnableAutomaticScanning)
                {
                    autoScanTimer.Start();
                }
                else
                {
                    autoScanTimer.Stop();
                }
                
                SetupDailySummaryTimer();
                if (settings.DailySummaryEnabled)
                {
                    dailySummaryTimer.Start();
                }
                else
                {
                    dailySummaryTimer.Stop();
                }
                
                // Update status
                UpdateStatusLabel();
                
                LogOutput("Settings updated successfully.");
            }
        }

        private void ViewJobIndexButton_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(settings.GoogleSheetJobIndexUrl))
            {
                System.Diagnostics.Process.Start(settings.GoogleSheetJobIndexUrl);
            }
            else
            {
                MessageBox.Show("No Job Index URL configured in settings.", 
                    "Missing URL", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DailySummaryButton_Click(object sender, EventArgs e)
        {
            dailySummaryButton.Enabled = false;
            try
            {
                Task.Run(async () => 
                {
                    try 
                    {
                        await GenerateDailyActivitySummary();
                    }
                    finally
                    {
                        // Use invoke to update UI from background thread
                        if (dailySummaryButton.InvokeRequired)
                        {
                            dailySummaryButton.Invoke(new Action(() => dailySummaryButton.Enabled = true));
                        }
                        else
                        {
                            dailySummaryButton.Enabled = true;
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                LogOutput($"Error generating daily summary: {ex.Message}");
                dailySummaryButton.Enabled = true;
            }
        }

        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    #endregion

    #region Settings Form

    public class SettingsForm : Form
    {
        private AppSettings _settings;
        private TabControl settingsTabControl;
        
        // Moraware connection controls
        private TextBox morawareUsernameTextBox;
        private TextBox morawarePasswordTextBox;
        private TextBox morawareDatabaseTextBox;
        
        // Google Sheets controls
        private TextBox googleSheetIdTextBox;
        private TextBox googleSheetScheduleNameTextBox;
        private TextBox googleSheetScheduleUrlTextBox;
        private TextBox googleSheetJobIndexNameTextBox;
        private TextBox googleSheetJobIndexUrlTextBox;
        private TextBox googleCredentialsPathTextBox;
        private TextBox jobIndexSheetRangeTextBox;
        private TextBox autofabSheetNameTextBox;
        
        // Automation controls
        private CheckBox enableAutoScanCheckBox;
        private NumericUpDown autoScanIntervalMinutes;
        private CheckBox enableDailySummaryCheckBox;
        private DateTimePicker dailySummaryTimePicker;
        
        // Buttons
        private Button saveButton;
        private Button cancelButton;
        private Button testMorawareButton;
        private Button testGoogleButton;
        private Button browseCredentialsButton;

        public AppSettings Settings { get { return _settings; } }

        public SettingsForm(AppSettings settings)
        {
            _settings = new AppSettings();
            // Make a deep copy of the settings
            _settings.MorawareUsername = settings.MorawareUsername;
            _settings.MorawarePassword = settings.MorawarePassword;
            _settings.MorawareDatabase = settings.MorawareDatabase;
            _settings.GoogleSheetScheduleName = settings.GoogleSheetScheduleName;
            _settings.GoogleSheetScheduleUrl = settings.GoogleSheetScheduleUrl;
            _settings.GoogleSheetJobIndexName = settings.GoogleSheetJobIndexName;
            _settings.GoogleSheetJobIndexUrl = settings.GoogleSheetJobIndexUrl;
            _settings.GoogleCredentialsPath = settings.GoogleCredentialsPath;
            _settings.GoogleSheetId = settings.GoogleSheetId;
            _settings.JobIndexSheetRange = settings.JobIndexSheetRange;
            _settings.AutofabSheetName = settings.AutofabSheetName;
            _settings.LastScannedJobNumber = settings.LastScannedJobNumber;
            _settings.EnableAutomaticScanning = settings.EnableAutomaticScanning;
            _settings.AutoScanIntervalMinutes = settings.AutoScanIntervalMinutes;
            _settings.LastScanDate = settings.LastScanDate;
            _settings.DailySummaryEnabled = settings.DailySummaryEnabled;
            _settings.DailySummaryTime = settings.DailySummaryTime;
            
            InitializeComponents();
            LoadSettingsToForm();
        }

        private void InitializeComponents()
        {
            this.Text = "Moralink Settings";
            this.Size = new Size(600, 500);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            
            // Main tab control
            settingsTabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Padding = new Point(10, 3)
            };
            
            // Create tabs
            var connectionTab = new TabPage("Moraware Connection");
            var googleSheetsTab = new TabPage("Google Sheets");
            var automationTab = new TabPage("Automation");
            
            // Build connection tab
            var connectionLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                RowCount = 4,
                ColumnCount = 2
            };
            
            connectionLayout.Controls.Add(new Label { Text = "Moraware Username:", Dock = DockStyle.Fill }, 0, 0);
            morawareUsernameTextBox = new TextBox { Dock = DockStyle.Fill };
            connectionLayout.Controls.Add(morawareUsernameTextBox, 1, 0);
            
            connectionLayout.Controls.Add(new Label { Text = "Moraware Password:", Dock = DockStyle.Fill }, 0, 1);
            morawarePasswordTextBox = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };
            connectionLayout.Controls.Add(morawarePasswordTextBox, 1, 1);
            
            connectionLayout.Controls.Add(new Label { Text = "Moraware Database:", Dock = DockStyle.Fill }, 0, 2);
            morawareDatabaseTextBox = new TextBox { Dock = DockStyle.Fill };
            connectionLayout.Controls.Add(morawareDatabaseTextBox, 1, 2);
            
            testMorawareButton = new Button
            {
                Text = "Test Moraware Connection",
                Dock = DockStyle.Fill
            };
            testMorawareButton.Click += TestMorawareButton_Click;
            
            connectionLayout.Controls.Add(testMorawareButton, 0, 3);
            connectionLayout.SetColumnSpan(testMorawareButton, 2);
            
            connectionTab.Controls.Add(connectionLayout);
            
            // Build Google Sheets tab
            var googleLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                RowCount = 8,
                ColumnCount = 3
            };
            
            googleLayout.Controls.Add(new Label { Text = "Google Credentials Path:", Dock = DockStyle.Fill }, 0, 0);
            googleCredentialsPathTextBox = new TextBox { Dock = DockStyle.Fill };
            googleLayout.Controls.Add(googleCredentialsPathTextBox, 1, 0);
            
            browseCredentialsButton = new Button { Text = "Browse...", Width = 80 };
            browseCredentialsButton.Click += BrowseCredentialsButton_Click;
            googleLayout.Controls.Add(browseCredentialsButton, 2, 0);
            
            googleLayout.Controls.Add(new Label { Text = "Google Sheet ID:", Dock = DockStyle.Fill }, 0, 1);
            googleSheetIdTextBox = new TextBox { Dock = DockStyle.Fill };
            googleLayout.Controls.Add(googleSheetIdTextBox, 1, 1);
            googleLayout.Controls.Add(new Label(), 2, 1);
            
            googleLayout.Controls.Add(new Label { Text = "Schedule Sheet Name:", Dock = DockStyle.Fill }, 0, 2);
            googleSheetScheduleNameTextBox = new TextBox { Dock = DockStyle.Fill };
            googleLayout.Controls.Add(googleSheetScheduleNameTextBox, 1, 2);
            googleLayout.Controls.Add(new Label(), 2, 2);
            
            googleLayout.Controls.Add(new Label { Text = "Schedule Sheet URL:", Dock = DockStyle.Fill }, 0, 3);
            googleSheetScheduleUrlTextBox = new TextBox { Dock = DockStyle.Fill };
            googleLayout.Controls.Add(googleSheetScheduleUrlTextBox, 1, 3);
            googleLayout.Controls.Add(new Label(), 2, 3);
            
            googleLayout.Controls.Add(new Label { Text = "Job Index Sheet Name:", Dock = DockStyle.Fill }, 0, 4);
            googleSheetJobIndexNameTextBox = new TextBox { Dock = DockStyle.Fill };
            googleLayout.Controls.Add(googleSheetJobIndexNameTextBox, 1, 4);
            googleLayout.Controls.Add(new Label(), 2, 4);
            
            googleLayout.Controls.Add(new Label { Text = "Job Index Sheet URL:", Dock = DockStyle.Fill }, 0, 5);
            googleSheetJobIndexUrlTextBox = new TextBox { Dock = DockStyle.Fill };
            googleLayout.Controls.Add(googleSheetJobIndexUrlTextBox, 1, 5);
            googleLayout.Controls.Add(new Label(), 2, 5);
            
            googleLayout.Controls.Add(new Label { Text = "Job Index Sheet Range:", Dock = DockStyle.Fill }, 0, 6);
            jobIndexSheetRangeTextBox = new TextBox { Dock = DockStyle.Fill };
            googleLayout.Controls.Add(jobIndexSheetRangeTextBox, 1, 6);
            googleLayout.Controls.Add(new Label(), 2, 6);
            
            googleLayout.Controls.Add(new Label { Text = "AUTOFAB Sheet Name:", Dock = DockStyle.Fill }, 0, 7);
            autofabSheetNameTextBox = new TextBox { Dock = DockStyle.Fill };
            googleLayout.Controls.Add(autofabSheetNameTextBox, 1, 7);
            
            testGoogleButton = new Button
            {
                Text = "Test Google Sheets Connection",
                Dock = DockStyle.Fill
            };
            testGoogleButton.Click += TestGoogleButton_Click;
            
            googleLayout.Controls.Add(testGoogleButton, 1, 8);
            googleLayout.SetColumnSpan(testGoogleButton, 2);
            
            googleSheetsTab.Controls.Add(googleLayout);
            
            // Build Automation tab
            var automationLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                RowCount = 6,
                ColumnCount = 2
            };
            
            // Job scanning section
            var scanningGroupBox = new GroupBox
            {
                Text = "Job Scanning",
                Dock = DockStyle.Fill
            };
            
            var scanningLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5),
                RowCount = 3,
                ColumnCount = 2
            };
            
            enableAutoScanCheckBox = new CheckBox
            {
                Text = "Enable Automatic Scanning",
                Dock = DockStyle.Fill
            };
            scanningLayout.Controls.Add(enableAutoScanCheckBox, 0, 0);
            scanningLayout.SetColumnSpan(enableAutoScanCheckBox, 2);
            
            scanningLayout.Controls.Add(new Label { Text = "Scan Interval (minutes):", Dock = DockStyle.Fill }, 0, 1);
            autoScanIntervalMinutes = new NumericUpDown
            {
                Minimum = 5,
                Maximum = 1440, // 24 hours
                Value = 120,
                Dock = DockStyle.Fill
            };
            scanningLayout.Controls.Add(autoScanIntervalMinutes, 1, 1);
            
            var lastScannedLabel = new Label
            {
                Text = $"Last Scanned Job: {_settings.LastScannedJobNumber}",
                Dock = DockStyle.Fill
            };
            scanningLayout.Controls.Add(lastScannedLabel, 0, 2);
            scanningLayout.SetColumnSpan(lastScannedLabel, 2);
            
            scanningGroupBox.Controls.Add(scanningLayout);
            automationLayout.Controls.Add(scanningGroupBox, 0, 0);
            automationLayout.SetColumnSpan(scanningGroupBox, 2);
            
            // Daily summary section
            var summaryGroupBox = new GroupBox
            {
                Text = "Daily Summary",
                Dock = DockStyle.Fill
            };
            
            var summaryLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5),
                RowCount = 2,
                ColumnCount = 2
            };
            
            enableDailySummaryCheckBox = new CheckBox
            {
                Text = "Enable Daily Summary Report",
                Dock = DockStyle.Fill
            };
            summaryLayout.Controls.Add(enableDailySummaryCheckBox, 0, 0);
            summaryLayout.SetColumnSpan(enableDailySummaryCheckBox, 2);
            
            summaryLayout.Controls.Add(new Label { Text = "Summary Time:", Dock = DockStyle.Fill }, 0, 1);
            dailySummaryTimePicker = new DateTimePicker
            {
                Format = DateTimePickerFormat.Time,
                ShowUpDown = true,
                Dock = DockStyle.Fill
            };
            summaryLayout.Controls.Add(dailySummaryTimePicker, 1, 1);
            
            summaryGroupBox.Controls.Add(summaryLayout);
            automationLayout.Controls.Add(summaryGroupBox, 0, 1);
            automationLayout.SetColumnSpan(summaryGroupBox, 2);
            
            automationTab.Controls.Add(automationLayout);
            
            // Build button panel
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 40,
                Padding = new Padding(0, 5, 10, 0)
            };
            
            cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Width = 80
            };
            buttonPanel.Controls.Add(cancelButton);
            
            saveButton = new Button
            {
                Text = "Save",
                DialogResult = DialogResult.OK,
                Width = 80
            };
            saveButton.Click += SaveButton_Click;
            buttonPanel.Controls.Add(saveButton);
            
            // Add tabs to tab control
            settingsTabControl.Controls.Add(connectionTab);
            settingsTabControl.Controls.Add(googleSheetsTab);
            settingsTabControl.Controls.Add(automationTab);
            
            // Add controls to form
            this.Controls.Add(settingsTabControl);
            this.Controls.Add(buttonPanel);
            
            this.AcceptButton = saveButton;
            this.CancelButton = cancelButton;
        }

        private void LoadSettingsToForm()
        {
            // Moraware connection
            morawareUsernameTextBox.Text = _settings.MorawareUsername;
            morawarePasswordTextBox.Text = _settings.MorawarePassword;
            morawareDatabaseTextBox.Text = _settings.MorawareDatabase;
            
            // Google Sheets
            googleSheetIdTextBox.Text = _settings.GoogleSheetId;
            googleSheetScheduleNameTextBox.Text = _settings.GoogleSheetScheduleName;
            googleSheetScheduleUrlTextBox.Text = _settings.GoogleSheetScheduleUrl;
            googleSheetJobIndexNameTextBox.Text = _settings.GoogleSheetJobIndexName;
            googleSheetJobIndexUrlTextBox.Text = _settings.GoogleSheetJobIndexUrl;
            googleCredentialsPathTextBox.Text = _settings.GoogleCredentialsPath;
            jobIndexSheetRangeTextBox.Text = _settings.JobIndexSheetRange;
            autofabSheetNameTextBox.Text = _settings.AutofabSheetName;
            
            // Automation
            enableAutoScanCheckBox.Checked = _settings.EnableAutomaticScanning;
            autoScanIntervalMinutes.Value = _settings.AutoScanIntervalMinutes;
            enableDailySummaryCheckBox.Checked = _settings.DailySummaryEnabled;
            
            if (!string.IsNullOrEmpty(_settings.DailySummaryTime))
            {
                try
                {
                    string[] timeParts = _settings.DailySummaryTime.Split(':');
                    int hour = int.Parse(timeParts[0]);
                    int minute = int.Parse(timeParts[1]);
                    
                    dailySummaryTimePicker.Value = new DateTime(
                        DateTime.Now.Year, 
                        DateTime.Now.Month, 
                        DateTime.Now.Day, 
                        hour, 
                        minute, 
                        0);
                }
                catch
                {
                    dailySummaryTimePicker.Value = DateTime.Parse("17:00");
                }
            }
            else
            {
                dailySummaryTimePicker.Value = DateTime.Parse("17:00");
            }
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            // Moraware connection
            _settings.MorawareUsername = morawareUsernameTextBox.Text;
            _settings.MorawarePassword = morawarePasswordTextBox.Text;
            _settings.MorawareDatabase = morawareDatabaseTextBox.Text;
            
            // Google Sheets
            _settings.GoogleSheetId = googleSheetIdTextBox.Text;
            _settings.GoogleSheetScheduleName = googleSheetScheduleNameTextBox.Text;
            _settings.GoogleSheetScheduleUrl = googleSheetScheduleUrlTextBox.Text;
            _settings.GoogleSheetJobIndexName = googleSheetJobIndexNameTextBox.Text;
            _settings.GoogleSheetJobIndexUrl = googleSheetJobIndexUrlTextBox.Text;
            _settings.GoogleCredentialsPath = googleCredentialsPathTextBox.Text;
            _settings.JobIndexSheetRange = jobIndexSheetRangeTextBox.Text;
            _settings.AutofabSheetName = autofabSheetNameTextBox.Text;
            
            // Automation
            _settings.EnableAutomaticScanning = enableAutoScanCheckBox.Checked;
            _settings.AutoScanIntervalMinutes = (int)autoScanIntervalMinutes.Value;
            _settings.DailySummaryEnabled = enableDailySummaryCheckBox.Checked;
            _settings.DailySummaryTime = dailySummaryTimePicker.Value.ToString("HH:mm");
        }

        private void TestMorawareButton_Click(object sender, EventArgs e)
        {
            testMorawareButton.Enabled = false;
            testMorawareButton.Text = "Testing...";
            
            Task.Run(() => 
            {
                try
                {
                    var database = morawareDatabaseTextBox.Text;
                    var username = morawareUsernameTextBox.Text;
                    var password = morawarePasswordTextBox.Text;
                    
                    if (string.IsNullOrWhiteSpace(database) ||
                        string.IsNullOrWhiteSpace(username) ||
                        string.IsNullOrWhiteSpace(password))
                    {
                        this.Invoke(new Action(() => 
                        {
                            MessageBox.Show("Please fill in all Moraware connection fields.",
                                "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }));
                        return;
                    }
                    
                    var connection = new Connection($"https://{database}.moraware.net/api.aspx", username, password);
                    connection.Connect();
                    
                    this.Invoke(new Action(() => 
                    {
                        MessageBox.Show("Successfully connected to Moraware!",
                            "Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }));
                    
                    connection.Disconnect();
                }
                catch (Exception ex)
                {
                    this.Invoke(new Action(() => 
                    {
                        MessageBox.Show($"Error connecting to Moraware: {ex.Message}",
                            "Connection Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
                finally
                {
                    this.Invoke(new Action(() => 
                    {
                        testMorawareButton.Enabled = true;
                        testMorawareButton.Text = "Test Moraware Connection";
                    }));
                }
            });
        }

        private void TestGoogleButton_Click(object sender, EventArgs e)
        {
            testGoogleButton.Enabled = false;
            testGoogleButton.Text = "Testing...";
            
            Task.Run(async () => 
            {
                try
                {
                    var credentialsPath = googleCredentialsPathTextBox.Text;
                    var sheetId = googleSheetIdTextBox.Text;
                    
                    if (string.IsNullOrWhiteSpace(credentialsPath) ||
                        string.IsNullOrWhiteSpace(sheetId))
                    {
                        this.Invoke(new Action(() => 
                        {
                            MessageBox.Show("Please fill in Google credentials path and sheet ID.",
                                "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }));
                        return;
                    }
                    
                    if (!File.Exists(credentialsPath))
                    {
                        this.Invoke(new Action(() => 
                        {
                            MessageBox.Show("Google credentials file not found at the specified path.",
                                "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }));
                        return;
                    }
                    
                    GoogleCredential credential;
                    using (var stream = new FileStream(credentialsPath, FileMode.Open, FileAccess.Read))
                    {
                        credential = GoogleCredential.FromStream(stream)
                            .CreateScoped(new[] { SheetsService.Scope.Spreadsheets });
                    }
                    
                    var service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = "Moralink Automation"
                    });
                    
                    // Try to get spreadsheet metadata
                    var request = service.Spreadsheets.Get(sheetId);
                    var response = await request.ExecuteAsync();
                    
                    this.Invoke(new Action(() => 
                    {
                        MessageBox.Show($"Successfully connected to Google Sheets!\nSpreadsheet Title: {response.Properties.Title}",
                            "Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }));
                }
                catch (Exception ex)
                {
                    this.Invoke(new Action(() => 
                    {
                        MessageBox.Show($"Error connecting to Google Sheets: {ex.Message}",
                            "Connection Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
                finally
                {
                    this.Invoke(new Action(() => 
                    {
                        testGoogleButton.Enabled = true;
                        testGoogleButton.Text = "Test Google Sheets Connection";
                    }));
                }
            });
        }

        private void BrowseCredentialsButton_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Select Google Credentials JSON File",
                Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                CheckFileExists = true
            };
            
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                googleCredentialsPathTextBox.Text = openFileDialog.FileName;
            }
        }
    }

    #endregion
}