using ClosedXML.Excel;
using Dish_ORAN_InitialSystem_GUI;
using DocumentFormat.OpenXml.EMMA;
using Microsoft.VisualBasic.ApplicationServices;
using Microsoft.VisualBasic.FileIO;
using Newtonsoft.Json.Linq;
using ORAN_Initial_System_GUI;
using Renci.SshNet;
using Renci.SshNet.Common;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.DirectoryServices;
using System.IO;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Sockets;
using System.Net.WebSockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Timers;
using System.Transactions;
using IOPath = System.IO.Path;  // Alias for System.IO.Path

namespace ORAN_Initial_System_GUI
{
    public partial class Form1 : Form
    {
        #region Fields and Constants

        int PCS_h, PCS_m, PCS_s = 0;
        int LOLO_h, LOLO_m, LOLO_s = 0;
        int slot;
        string Settings = @"C:\Carrier Oran Settings";
        string LOLOIP = @"C:\Carrier Oran Settings\LOLO IP.txt";
        string PCSIP = @"C:\Carrier Oran Settings\PCS IP.txt";
        string LOLORuId = @"C:\Carrier Oran Settings\LOLO RU ID.txt";
        string PCSRuId = @"C:\Carrier Oran Settings\PCS RU ID.txt";
        bool exitisclicked = false;
        public Setting settings = new Setting();
        VZ_LOLO vZ_LOLO = new VZ_LOLO();
        VZ_PCS vZ_PCS = new VZ_PCS();

        #region RegexHelper    
        public static class RegexHelper
        {
            private static readonly Regex fpgaRegex = new Regex(
                @"FpgaTemp\|\s+([\d\.]+)\|\s+([\d\.]+)\|\s+([\d\.]+)",
                RegexOptions.IgnoreCase | RegexOptions.Compiled
            );

            private static readonly Regex boardRegex = new Regex(
                @"BoardTemp\|\s+([\d\.]+)\|\s+([\d\.]+)\|\s+([\d\.]+)",
                RegexOptions.IgnoreCase | RegexOptions.Compiled
            );


            private static readonly Regex rlRegex = new Regex(
                @"ReturnLoss\s*\[\s*dB\]\s*:\s*\[\s*([\d\.\-]+)\]\s*\[\s*([\d\.\-]+)\]\s*\[\s*([\d\.\-]+)\]\s*\[\s*([\d\.\-]+)\]\s*\[\s*([\d\.\-]+)\]\s*\[\s*([\d\.\-]+)\]\s*\[\s*([\d\.\-]+)\]\s*\[\s*([\d\.\-]+)\]",
                RegexOptions.IgnoreCase | RegexOptions.Compiled
            );

            // New "Ret Loss| ... dB|" table format for CarrierB
            private static readonly Regex rlTableRegex = new Regex(
                @"Ret Loss\|\s*([\d\.\-]+) dB\|\s*([\d\.\-]+) dB\|\s*([\d\.\-]+) dB\|\s*([\d\.\-]+) dB\|\s*([\d\.\-]+) dB\|\s*([\d\.\-]+) dB\|\s*([\d\.\-]+) dB\|\s*([\d\.\-]+) dB\|",
                RegexOptions.IgnoreCase | RegexOptions.Compiled
            );

            // alarm regex
            private static readonly Regex alarmRegex = new Regex(
                @"\]\s*(.*?)\s*:",
                RegexOptions.IgnoreCase | RegexOptions.Compiled
            );

            // Alarm ID 
            private static readonly Regex idRegex = new Regex(
                @":\s+(\d+)\s+\(\s*(\d+)\)",
                RegexOptions.IgnoreCase | RegexOptions.Compiled
            );

            public static Match MatchFpga(string input) => fpgaRegex.Match(input);
            public static Match MatchBoard(string input) => boardRegex.Match(input);
            public static Match MatchReturnLoss(string input) => rlRegex.Match(input);
            //CarrierB
            public static Match MatchReturnLossTable(string input) => rlTableRegex.Match(input);
            public static Match MatchAlarm(string input) => alarmRegex.Match(input);
            public static Match MatchId(string input) => idRegex.Match(input);
        }


        #endregion

        enum DataRow : int
        {
            TxAntPow = 1,
            RSSI = 2,
            RET_Loss = 3
        }
        Dictionary<string, string> directory = new()
        {
            {"PCS", @"T:\Acme Test Logs\5G RU ORAN\CarrierA PCS\" },
            {"LOLO", @"T:\Acme Test Logs\5G RU ORAN\CarrierA LOLO\" },
            {"FAT LOLO", @"T:\Acme Test Logs\5G RU ORAN\CarrierA FAT LOLO\" }
        };

        #endregion Fields and Constants

        #region Helper Methods

        private static void setGridValue(DataGridView grid, DataRow forRow, int forPath, String value, Boolean isGreen)
        {
            grid.Rows[(int)forRow].Cells[forPath].Value = value;
            grid.Rows[(int)forRow].Cells[forPath].Style.BackColor = isGreen ? Color.LightGreen : Color.LightPink;
        }
        #endregion

        #region Constructor

        public Form1()
        {
            settings.getJoke = false;
            InitializeComponent();
            MapT_Drive t_Drive = new MapT_Drive();
            t_Drive.EnsureTDriveMapped();
            dataGridView1.Show();
            dataGridView1.Rows.Add("Antenna ID", 1, 2, 3, 4, 5, 6, 7, 8);
            dataGridView1.Rows.Add("TxAntPow");
            dataGridView1.Rows.Add("RSSI");
            dataGridView1.Rows.Add("RET Loss");
            dataGridView1.DefaultCellStyle = new DataGridViewCellStyle() { Alignment = DataGridViewContentAlignment.MiddleCenter };
            this.FormClosing += new FormClosingEventHandler(MainForm_FormClosing);
            dataGridView2.Show();
            dataGridView2.Rows.Add("Antenna ID", 1, 2, 3, 4, 5, 6, 7, 8);
            dataGridView2.Rows.Add("TxAntPow");
            dataGridView2.Rows.Add("RSSI");
            dataGridView2.Rows.Add("RET Loss");
            dataGridView2.DefaultCellStyle = new DataGridViewCellStyle() { Alignment = DataGridViewContentAlignment.MiddleCenter };

            if (!Directory.Exists(Settings) || !File.Exists(LOLOIP)) // Check if IP addres folder and Text File Exist
            {
                Directory.CreateDirectory(Settings);                             // Create IP address Folder
                using (FileStream fstri = File.Create(LOLOIP))         // Create IP address Text File and Write the standard IP for Tri Band REDACTED_IP
                {
                    // Add some text to file
                    Byte[] ipaddrestri = new UTF8Encoding(true).GetBytes("REDACTED_IP");
                    fstri.Write(ipaddrestri, 0, ipaddrestri.Length);
                }
                using (FileStream fstriruid = File.Create(LOLORuId))
                {
                    // Add some text to file
                    Byte[] ruidtri = new UTF8Encoding(true).GetBytes("702");
                    fstriruid.Write(ruidtri, 0, ruidtri.Length);
                }
            }

            if (!Directory.Exists(Settings) || !File.Exists(PCSIP))// Check if IP addres folder and Text File Exist
            {
                Directory.CreateDirectory(Settings);                             // Create IP address Folder

                using (FileStream fsdual = File.Create(PCSIP))  // Create IP address Text File and Write the standard IP for Dual Band REDACTED_IP
                {
                    // Add some text to file
                    Byte[] ipaddresdual = new UTF8Encoding(true).GetBytes("REDACTED_IP");
                    fsdual.Write(ipaddresdual, 0, ipaddresdual.Length);
                }

                using (FileStream fsdualruid = File.Create(PCSRuId))
                {
                    // Add some text to file
                    Byte[] ruiddual = new UTF8Encoding(true).GetBytes("702");
                    fsdualruid.Write(ruiddual, 0, ruiddual.Length);
                }
            }

            LOLOIPLabel.Text = File.ReadAllText(LOLOIP);
            LOLORUIDLabel.Text = File.ReadAllText(LOLORuId);
            PCSIPLabel.Text = File.ReadAllText(PCSIP);
            PCSRUIDLabel.Text = File.ReadAllText(PCSRuId);
        }

        #endregion Constructor

        #region FTP Connection and File Download

        private async Task<DataTable> GetDataAsync(string serialNumber)
        {
            string ftpServer = "ftp://sftp.example.com";
            string ftpUser = "REDACTED_USER";
            string ftpPassword = "REDACTED_PASSWORD";
            string remoteDirPath = "/receiving";
            string localDirPath = @"C:\OLP File";

            if (!Directory.Exists(localDirPath))
            {
                Directory.CreateDirectory(localDirPath);
            }

            if (string.IsNullOrEmpty(serialNumber))
            {
                MessageBox.Show("No serial number entered.");
                return null;
            }

            //logBox.Invoke(() => {
            //    logBox.SelectionColor = Color.DarkBlue;
            //    logBox.AppendText($"[{DateTime.Now}] Checking local Excel for serial number {serialNumber}...\n");
            //    logBox.SelectionColor = Color.Black;
            //});

            string latestLocalFile = await GetLatestExcelFileAsync(localDirPath);
            DataTable localTable = null;

            if (!string.IsNullOrEmpty(latestLocalFile))
            {
                localTable = LoadExcelToDataTable(latestLocalFile);
                if (localTable != null)
                {
                    var match = localTable.AsEnumerable()
                        .FirstOrDefault(row => row["SERIALNBR"].ToString().Trim().Equals(serialNumber, StringComparison.OrdinalIgnoreCase));

                    if (match != null)
                    {
                        //logBox.Invoke(() => {
                        //    logBox.SelectionColor = Color.Green;
                        //    logBox.AppendText($"[{DateTime.Now}] Serial number found in local file.\n");
                        //    logBox.SelectionColor = Color.Black;
                        //});
                        return localTable; //  Serial found locally
                    }
                }
            }

            // If not found, download latest file from FTP
            //logBox.Invoke(() => {
            //    logBox.SelectionColor = Color.Orange;
            //    logBox.AppendText($"[{DateTime.Now}] Serial not found. Downloading latest file from FTP...\n");
            //    logBox.SelectionColor = Color.Black;
            //});

            try
            {
                Directory.CreateDirectory(localDirPath);
                await DownloadLatestFileFromFTPAsync(ftpServer, ftpUser, ftpPassword, remoteDirPath, localDirPath/*, logBox*/);
            }
            catch (Exception ex)
            {
                //logBox.Invoke(() => {
                //    logBox.SelectionColor = Color.Red;
                //    logBox.AppendText($"[{DateTime.Now}] FTP error: {ex.Message}\n");
                //    logBox.SelectionColor = Color.Black;
                //});
                return null;
            }

            string latestFileAfterDownload = await GetLatestExcelFileAsync(localDirPath);
            if (string.IsNullOrEmpty(latestFileAfterDownload))
            {
                //logBox.Invoke(() => {
                //    logBox.SelectionColor = Color.Red;
                //    logBox.AppendText($"[{DateTime.Now}] No Excel files found in {localDirPath}\n");
                //    logBox.SelectionColor = Color.Black;
                //});
                return null;
            }

            var newTable = LoadExcelToDataTable(latestFileAfterDownload);

            //logBox.Invoke(() => {
            //    logBox.SelectionColor = Color.Green;
            //    logBox.AppendText($"[{DateTime.Now}] Loaded Excel: {IOPath.GetFileName(latestFileAfterDownload)}\n");
            //    logBox.SelectionColor = Color.Black;
            //});

            return newTable;
        }

        private DataTable LoadExcelToDataTable(string excelFilePath)
        {
            try
            {
                using (var workbook = new XLWorkbook(excelFilePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    DataTable table = new DataTable();
                    bool isFirstRow = true;

                    foreach (var row in worksheet.RowsUsed())
                    {
                        if (isFirstRow)
                        {
                            foreach (var cell in row.Cells())
                                table.Columns.Add(cell.Value.ToString().Trim());
                            isFirstRow = false;
                        }
                        else
                        {
                            table.Rows.Add(row.Cells().Select(c => c.Value.ToString().Trim()).ToArray());
                        }
                    }
                    return table;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Excel load error: " + ex.Message);
                return null;
            }
        }


        public static async Task DownloadLatestFileFromFTPAsync(
    string ftpServer, string ftpUser, string ftpPassword,
    string remoteDirPath, string localDirPath/*, RichTextBox logBox*/)
        {
            await Task.Run(() => {
                try
                {
                    if (!ftpServer.EndsWith("/")) ftpServer += "/";
                    if (!remoteDirPath.EndsWith("/")) remoteDirPath += "/";
                    string fullFtpDir = ftpServer + remoteDirPath;

                    // Regex for Acme_Receiving_20250808_224831_.xlsx
                    Regex filePattern = new Regex(
                        @"Acme_Receiving_(\d{8})_(\d{6})_\.xlsx",
                        RegexOptions.IgnoreCase
                    );

                    // Step 1: Get the list of files (fast - only file names)
                    FtpWebRequest listRequest = (FtpWebRequest)WebRequest.Create(fullFtpDir);
                    listRequest.Method = WebRequestMethods.Ftp.ListDirectory;
                    listRequest.Credentials = new NetworkCredential(ftpUser, ftpPassword);

                    List<string> files = new List<string>();

                    using (FtpWebResponse listResponse = (FtpWebResponse)listRequest.GetResponse())
                    using (StreamReader reader = new StreamReader(listResponse.GetResponseStream()))
                    {
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            string fileName = IOPath.GetFileName(line.Trim());
                            if (!string.IsNullOrEmpty(fileName))
                                files.Add(fileName);
                        }
                    }

                    if (files.Count == 0)
                    {
                        //logBox.Invoke(() => logBox.AppendText($"[{DateTime.Now}] No files found in FTP directory.\n"));
                        return;
                    }

                    // Step 2: Find latest file based on name timestamp
                    string latestFile = null;
                    DateTime latestDate = DateTime.MinValue;

                    foreach (string file in files)
                    {
                        Match match = filePattern.Match(file);
                        if (match.Success)
                        {
                            string datePart = match.Groups[1].Value; // YYYYMMDD
                            string timePart = match.Groups[2].Value; // HHMMSS

                            if (DateTime.TryParseExact(
                                datePart + timePart,
                                "yyyyMMddHHmmss",
                                null,
                                System.Globalization.DateTimeStyles.None,
                                out DateTime fileDate))
                            {
                                if (fileDate > latestDate)
                                {
                                    latestDate = fileDate;
                                    latestFile = file;
                                }
                            }
                        }
                    }

                    if (latestFile == null)
                    {
                        //logBox.Invoke(() => logBox.AppendText($"[{DateTime.Now}] No matching files found by pattern.\n"));
                        return;
                    }

                    // Step 3: Download the latest file
                    string latestFileUrl = fullFtpDir + latestFile;
                    string localFilePath = IOPath.Combine(localDirPath, latestFile);

                    Directory.CreateDirectory(localDirPath);

                    FtpWebRequest downloadRequest = (FtpWebRequest)WebRequest.Create(latestFileUrl);
                    downloadRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                    downloadRequest.Credentials = new NetworkCredential(ftpUser, ftpPassword);

                    using (FtpWebResponse downloadResponse = (FtpWebResponse)downloadRequest.GetResponse())
                    using (Stream responseStream = downloadResponse.GetResponseStream())
                    using (FileStream localFileStream = new FileStream(localFilePath, FileMode.Create))
                    {
                        responseStream.CopyTo(localFileStream);
                        //logBox.Invoke(() => logBox.AppendText($"[{DateTime.Now}] Downloaded latest file: {latestFile}\n"));
                    }
                }
                catch (Exception ex)
                {
                    //logBox.Invoke(() => logBox.AppendText($"[{DateTime.Now}] FTP Download Error: {ex.Message}\n"));
                }
            });
        }


        private static Task<string> GetLatestExcelFileAsync(string folderPath)
        {
            return Task.Run(() => {
                var directory = new DirectoryInfo(folderPath);
                var file = directory.GetFiles("*.xlsx")
                                    .OrderByDescending(f => f.LastWriteTime)
                                    .FirstOrDefault();
                return file?.FullName;
            });
        }

        #endregion

        #region Model Mapping

        private string MapPartNumberToModel(string partNumber) => partNumber switch
        {
            "SFG-ARR57201VZ" => "FAT LOLO",
            "SFG-ARR27201VZ" => "LOLO",
            "SFG-ARR26301VZ" => "PCS",
            _ => "UNKNOWN"
        };

        #endregion Model Mapping

        #region Form Events

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Check if the user is trying to close the form

            DialogResult result = MessageBox.Show("Are you sure you want to exit?", "Confirm Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            // Check if the user clicked "Yes"
            if (result == DialogResult.No)
            {
                // Close the form
                e.Cancel = true;
            }
            // If the user clicked "No", do nothing and keep the form open
            else
            {
                e.Cancel = false;
            }

        }

        #endregion Form Events

        #region PCS

        private void PCSRefreshButton_Click(object sender, EventArgs e)
        {
            PCSIPLabel.Text = File.ReadAllText(PCSIP);
            PCSRUIDLabel.Text = File.ReadAllText(PCSRuId);
        }
        private async void HandlePCSSerialNumber()
        {
            string model = string.Empty;
            if (scannedSN_PCS.Text != null && scannedSN_PCS.Text.Length == 10)
            {
                string serialNumber = scannedSN_PCS.Text.ToString();
                var excelTable = await GetDataAsync(serialNumber);
                var match = excelTable.AsEnumerable()
             .FirstOrDefault(row => row["SERIALNBR"].ToString().Trim()
                 .Equals(serialNumber, StringComparison.OrdinalIgnoreCase));
                if (match != null)
                {
                    string partNumber = match["PARTNBR"].ToString().Trim();
                    model = MapPartNumberToModel(partNumber);
                    modeltextslot1.Text = model;
                    if (model != "PCS")
                    {
                        MessageBox.Show("In Correct Model..!" + " This is Model: " + model);
                    }
                }
            }
        }
        private void scannedSN_PCS_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                HandlePCSSerialNumber();
            }
        }
        private async void startButton_PCS_Click(object sender, EventArgs e)
        {
            if (scannedSN_PCS.Text != null && scannedSN_PCS.Text.Length == 10)
            {
                clearButtonPCS_Click(this, e);
                string serialNumber = scannedSN_PCS.Text.ToString();
                var excelTable = await GetDataAsync(serialNumber);

                var match = excelTable.AsEnumerable()
             .FirstOrDefault(row => row["SERIALNBR"].ToString().Trim()
                 .Equals(serialNumber, StringComparison.OrdinalIgnoreCase));

                if (match != null)
                {
                    string partNumber = match["PARTNBR"].ToString().Trim();
                    string model = MapPartNumberToModel(partNumber);
                    modeltextslot1.Text = model;
                    if (model == "PCS")
                    {
                        Thread x = new(new ThreadStart(BeginPCSTest)) { IsBackground = true };
                        x.Start();
                        startButton_PCS.Enabled = false;
                        clearButtonPCS.Enabled = false;
                        scannedSN_PCS.Enabled = false;
                        PCS_TimerLabel.Text = "00:00:00";

                    }
                    else
                    {
                        MessageBox.Show("Incorrect Model..!" + " This is Model: " + model);
                    }
                }

            }
            else
            {
                MessageBox.Show("Invalid serial number. Test not started");
            }

        }
        private async void clearButtonPCS_Click(object sender, EventArgs e)
        {
            PCS_TimerLabel.Text = "00:00:00";
            PCS_s = 0;
            PCS_m = 0;
            PCS_h = 0;
            testLog_PCS.Text = string.Empty;
            for (int i = 1; i <= 3; i++)
            {
                for (int j = 1; j <= 8; j++)
                {
                    dataGridView1.Rows[i].Cells[j].Value = "";
                    dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.White;
                }
            }


        }
        private void testLog_PCS_TextChanged(object sender, EventArgs e)
        {
            testLog_PCS.SelectionStart = testLog_PCS.Text.Length;
            testLog_PCS.ScrollToCaret();
        }
        private void BeginPCSTest()
        {
            PCS_s = 0;
            PCS_m = 0;

            System.Timers.Timer PCS_timer = new System.Timers.Timer
            {
                Interval = 1000
            };
            PCS_timer.Elapsed += OnTimeEvent;
            void OnTimeEvent(object? sender, ElapsedEventArgs e)
            {
                Invoke(new Action(() => {
                    PCS_s += 1;
                    if (PCS_s == 60)
                    {
                        PCS_s = 0;
                        PCS_m += 1;
                    }
                    if (PCS_m == 60)
                    {
                        PCS_m = 0;
                        PCS_h += 1;
                    }
                    PCS_TimerLabel.Text = string.Format("{0}:{1}:{2}", PCS_h.ToString().ToString().PadLeft(2, '0'), PCS_m.ToString().ToString().PadLeft(2, '0'), PCS_s.ToString().ToString().PadLeft(2, '0'));
                }));
            }
            PCS_timer.Start();
            StartTest(PCSIPLabel.Text, PCSRUIDLabel.Text, dataGridView1, testLog_PCS, scannedSN_PCS, "PCS", PCS_timer, PCS_TimerLabel);
            this.Invoke(new MethodInvoker(delegate {
                startButton_PCS.Enabled = true;
                clearButtonPCS.Enabled = true;
                scannedSN_PCS.Enabled = true;
            }));
            PCS_timer.Stop();
        }
        #endregion PCS

        #region LOLO
        private async void HandleLOLOSerialNumber()
        {
            string model = string.Empty;
            if (scannedSN_LOLO.Text != null && scannedSN_LOLO.Text.Length == 10)
            {
                string serialNumber = scannedSN_LOLO.Text.ToString();
                var excelTable = await GetDataAsync(serialNumber);
                var match = excelTable.AsEnumerable()
             .FirstOrDefault(row => row["SERIALNBR"].ToString().Trim()
                 .Equals(serialNumber, StringComparison.OrdinalIgnoreCase));
                if (match != null)
                {
                    string partNumber = match["PARTNBR"].ToString().Trim();
                    model = MapPartNumberToModel(partNumber);
                    modeltextslot2.Text = model;
                    if (model != "LOLO" && model != "FAT LOLO")
                    {
                        MessageBox.Show("InCorrect Model..!" + " This is Model: " + model);
                    }
                    if (model == "LOLO" && !LOLORUIDLabel.Text.Contains("7"))
                    {
                        LOLORUIDLabel.Text = "700";
                    }
                    else if (model == "FAT LOLO" && !LOLORUIDLabel.Text.Contains("8"))
                    {
                        LOLORUIDLabel.Text = "800";
                    }
                }
            }
        }
        private void scannedSN_LOLO_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                HandleLOLOSerialNumber();
            }
        }
        private void LOLORefreshButton_Click(object sender, EventArgs e)
        {
            LOLOIPLabel.Text = File.ReadAllText(LOLOIP);
            LOLORUIDLabel.Text = File.ReadAllText(LOLORuId);
        }
        private async void startButton_LOLO_Click(object sender, EventArgs e)
        {
            if (scannedSN_LOLO.Text != null && scannedSN_LOLO.Text.Length == 10)
            {
                clearButtonLOLO_Click(this, e);
                string serialNumber = scannedSN_LOLO.Text.ToString();
                var excelTable = await GetDataAsync(serialNumber);

                var match = excelTable.AsEnumerable()
               .FirstOrDefault(row => row["SERIALNBR"].ToString().Trim()
                   .Equals(serialNumber, StringComparison.OrdinalIgnoreCase));

                if (match != null)
                {
                    string partNumber = match["PARTNBR"].ToString().Trim();
                    string model = MapPartNumberToModel(partNumber);
                    modeltextslot2.Text = model;
                    if (model == "LOLO" || model == "FAT LOLO")
                    {
                        Thread x = new(new ThreadStart(BeginLOLOTest)) { IsBackground = true };
                        x.Start();
                        startButton_LOLO.Enabled = false;
                        clearButtonLOLO.Enabled = false;
                        scannedSN_LOLO.Enabled = false;
                        LOLO_TimerLabel.Text = "00:00:00";

                    }
                    else
                    {
                        MessageBox.Show("Incorrect Model..!" + " This is Model: " + model);
                    }

                }


            }
            else
            {
                MessageBox.Show("Invalid serial number. Test not started");
            }


        }
        private void BeginLOLOTest()
        {
            LOLO_h = 0;
            LOLO_s = 0;
            LOLO_m = 0;
            using System.Timers.Timer LOLO_timer = new System.Timers.Timer
            {
                Interval = 1000
            };
            LOLO_timer.Elapsed += OnTimeEvent;
            void OnTimeEvent(object? sender, ElapsedEventArgs e)
            {
                Invoke(new Action(() => {
                    LOLO_s += 1;
                    if (LOLO_s == 60)
                    {
                        LOLO_s = 0;
                        LOLO_m += 1;
                    }
                    if (LOLO_m == 60)
                    {
                        LOLO_m = 0;
                        LOLO_h += 1;
                    }
                    LOLO_TimerLabel.Text = string.Format("{0}:{1}:{2}", LOLO_h.ToString().ToString().PadLeft(2, '0'), LOLO_m.ToString().ToString().PadLeft(2, '0'), LOLO_s.ToString().ToString().PadLeft(2, '0'));
                }));
            }
            string loloModel = modeltextslot2.Text;
            /*if (loloModel == "LOLO") {
                string originalText = LOLORUIDLabel.Text;
                // Replace the first character with '7' and keep the rest of the string
                string updatedText = "7" + originalText.Substring(1);
                LOLORUIDLabel.Text = updatedText;
            } else if (loloModel == "FAT LOLO") {
                string originalText = LOLORUIDLabel.Text;
                // Replace the first character with '7' and keep the rest of the string
                string updatedText = "8" + originalText.Substring(1);
                LOLORUIDLabel.Text = updatedText;
            }*/
            LOLO_timer.Start();
            StartTest(LOLOIPLabel.Text, LOLORUIDLabel.Text, dataGridView2, testLog_LOLO, scannedSN_LOLO, loloModel, LOLO_timer, LOLO_TimerLabel);
            this.Invoke(new MethodInvoker(delegate {
                startButton_LOLO.Enabled = true;
                clearButtonLOLO.Enabled = true;
                scannedSN_LOLO.Enabled = true;
            }));
            LOLO_timer.Stop();
        }
        private void testLog_LOLO_TextChanged(object sender, EventArgs e)
        {
            testLog_LOLO.SelectionStart = testLog_LOLO.Text.Length;
            testLog_LOLO.ScrollToCaret();
        }
        private void clearButtonLOLO_Click(object sender, EventArgs e)
        {
            LOLO_TimerLabel.Text = "00:00:00";
            LOLO_s = 0;
            LOLO_m = 0;
            LOLO_h = 0;
            testLog_LOLO.Text = string.Empty;
            for (int i = 1; i <= 3; i++)
            {
                for (int j = 1; j <= 8; j++)
                {
                    dataGridView2.Rows[i].Cells[j].Value = "";
                    dataGridView2.Rows[i].Cells[j].Style.BackColor = Color.White;
                }
            }
        }
        #endregion LOLO

        #region Utility Methods

        private string RemoveAnsiCodes(string text)
        {
            return Regex.Replace(text, @"\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])", "");
        }

        private static string EscapeSingleQuotesForShell(string path)
        {
            // POSIX-safe escape: ' -> '"'"'
            return path.Replace("'", "'\"'\"'");
        }
        #endregion Utility Methods

        #region Log and Parsing Methods

        public List<string> LogBuilder(string serialNumber, string Model)
        {
            List<string> list = new List<string>();
            DateTime now = DateTime.Now;
            string date = now.ToString("yyMMdd");
            string dir = @"C:\Log\";
            string dir2 = string.Empty;
            if (Model == "PCS")
            {
                dir += @"CarrierA_PCS\";
            }
            else if (Model == "LOLO")
            {
                dir += @"CarrierA_Lo_Lo\";
            }
            else if (Model == "FAT LOLO")
            {
                dir += @"CarrierA_Lo_Lo_XL\";
            }
            Directory.CreateDirectory(dir);
            string logfile = serialNumber;
            switch (Model)
            {
                case "PCS":
                    this.Invoke(new MethodInvoker(delegate {
                        if (radioInitial_PCS.Checked)
                        {
                            logfile += "_Initial_";
                            dir2 = @"Initial\";
                        }
                        else if (radioSystem_PCS.Checked)
                        {
                            logfile += "_System_";
                            dir2 = @"System Test\";
                        }
                    }));
                    break;
                case "LOLO":
                    this.Invoke(new MethodInvoker(delegate {
                        if (radioInitial_LOLO.Checked)
                        {
                            logfile += "_Initial_";
                            dir2 = @"Initial\";
                        }
                        else if (radioSystem_LOLO.Checked)
                        {
                            logfile += "_System_";
                            dir2 = @"System Test\";
                        }
                    }));
                    break;
                case "FAT LOLO":
                    this.Invoke(new MethodInvoker(delegate {
                        if (radioInitial_LOLO.Checked)
                        {
                            logfile += "_Initial_";
                            dir2 = @"Initial\";
                        }
                        else if (radioSystem_LOLO.Checked)
                        {
                            logfile += "_System_";
                            dir2 = @"System Test\";
                        }
                    }));
                    break;
            }
            list.Add(dir + logfile + date + ".txt");
            dir = directory[Model] + dir2;
            list.Add(dir + logfile + date + ".txt");
            return list;
        }

        public bool ParseAntPowReadings(DataGridView grid, string input, LogHandler logger, string Model)
        {
            var lines = input.Split("\r\n");
            bool result = true;
            double low = 43.5;
            double high = 46.5;
            foreach (var line in lines)
            {
                if (line.Contains("TxAntSum"))
                {
                    string[] values = line.Split(new String[] { " ", "|" }, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 1; i < values.Length; i++)
                    {
                        if (double.Parse(values[i]) < low || double.Parse(values[i]) > high)
                        {
                            logger.tfailed.Add(new TestFailed { TestName = "Tx Power : Path " + i, Value = values[i], Result = "FAIL", ErrorCodes = "TT053" });
                            result = false;
                        }
                        else
                        {
                            result = true;
                        }
                        setGridValue(grid, DataRow.TxAntPow, i, values[i], result);
                    }
                }
                if (line.Contains("RxAntFa00"))
                {
                    string[] values = line.Split(new String[] { " ", "|" }, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 1; i < values.Length; i++)
                    {
                        if (double.Parse(values[i]) > -96 || double.Parse(values[i]) < -107)
                        {
                            if (double.Parse(values[i]) > -30 || double.Parse(values[i]) < -120)
                            {
                                logger.tfailed.Add(new TestFailed { TestName = "RSSI : Path " + i, Value = values[i], Result = "FAIL", ErrorCodes = "TT045" });
                                result = false;
                            }

                        }
                        else
                        {
                            result = true;
                        }
                        setGridValue(grid, DataRow.RSSI, i, values[i], result);
                    }
                }
            }
            return result;
        }
        private static string StringCleaner(string input, string startpoint, string endpoint)
        {
            string result = "";
            bool isWriting = false;
            string[] lines = input.Split("\r\n");

            foreach (string line in lines)
            {
                if (line.Contains(startpoint))
                {
                    isWriting = true;
                }
                else if (line.Contains(endpoint))
                {
                    isWriting = false;
                }
                if (isWriting)
                {
                    result += line + "\r\n";
                }
            }
            return result;
        }
        public (bool, bool) ParseAlarms(string input, RichTextBox testLog, LogHandler logger)
        {
            bool alarmNotFound = true;
            bool vswrDetected = false;
            var lines = input.Split("\r\n");
            string primary = "PRIMARY ";
            foreach (var line in lines)
            {
                if (line.Length > 70)
                {
                    if (line.Contains("VswrFail(MJ)")) { vswrDetected = true; }
                    if (!line.Contains("UDA")
                        && !line.Contains("ShutDown")
                        && !line.Contains("GroupTxWithAlarm")
                        && !line.Contains("HighPimLevel")
                        && !line.Contains("OptRxLOS")
                        && !line.Contains("RxOverflowStep")
                        && !line.Contains("GroupTxShutdown")
                        && !line.Contains("LowGainSymptom")
                        && !line.Contains("RssiImbalance"))
                    {
                        var values = line.Split(new String[] { "|", " " }, StringSplitOptions.RemoveEmptyEntries);
                        switch (values.Length)
                        {
                            case 11:
                                {
                                    this.Invoke(new MethodInvoker(delegate {
                                        testLog.SelectionColor = Color.Red;
                                        testLog.AppendText(values[2].Replace("\u001b[33m", "").Replace("\u001b[31m", "") + " has been detected on Path ID" + values[4] + "\n");
                                        testLog.SelectionColor = Color.White;
                                    }));
                                    logger.tfailed.Add(new TestFailed { TestName = /*primary +*/"Alarm : ", Value = RemoveAnsiCodes(values[2]) + " Path ID: " + values[4], Result = "FAIL", ErrorCodes = "TT045" });
                                    //primary = string.Empty;
                                    alarmNotFound = false;
                                }
                                break;
                            case 8:
                                {
                                    this.Invoke(new MethodInvoker(delegate {
                                        testLog.SelectionColor = Color.Red;
                                        testLog.AppendText(values[1].Replace("\u001b[33m", "").Replace("\u001b[31m", "") + " has been detected on " + values[7] + "\n");
                                        testLog.SelectionColor = Color.White;
                                    }));
                                    logger.tfailed.Add(new TestFailed { TestName = /*primary +*/ "Alarm : ", Value = RemoveAnsiCodes(values[1]) + " Path ID: " + values[7], Result = "FAIL", ErrorCodes = "TT045" });
                                    //primary = string.Empty;
                                    alarmNotFound = false;
                                }
                                break;
                            case 10:
                                {
                                    this.Invoke(new MethodInvoker(delegate {
                                        testLog.SelectionColor = Color.Red;
                                        testLog.AppendText(values[1].Replace("\u001b[33m", "").Replace("\u001b[31m", "") + " has been detected on " + values[3] + "\n");
                                        testLog.SelectionColor = Color.White;
                                    }));
                                    logger.tfailed.Add(new TestFailed { TestName = /*primary +*/ "Alarm : ", Value = RemoveAnsiCodes(values[1]) + " Path ID: " + values[3], Result = "FAIL", ErrorCodes = "TT045" });
                                    //primary = string.Empty;
                                    alarmNotFound = false;
                                }
                                break;
                            case 9:
                                this.Invoke(new MethodInvoker(delegate {
                                    testLog.SelectionColor = Color.Red;
                                    testLog.AppendText(values[2].Replace("\u001b[33m", "").Replace("\u001b[31m", "") + " has been detected on " + values[8] + "\n");
                                    testLog.SelectionColor = Color.White;
                                }));
                                logger.tfailed.Add(new TestFailed { TestName = /*primary +*/ "Alarm : ", Value = RemoveAnsiCodes(values[2]) + " Path ID: " + values[8], Result = "FAIL", ErrorCodes = "TT045" });
                                //primary = string.Empty;
                                alarmNotFound = false;
                                break;
                            default:
                                {
                                    this.Invoke(new MethodInvoker(delegate {
                                        testLog.SelectionColor = Color.Red;
                                        testLog.AppendText("\nUnknown alarm detected. Send log file to Engineer for evaluation\n");
                                        testLog.SelectionColor = Color.White;
                                    }));
                                }
                                break;
                        }
                    }
                }
            }
            return (alarmNotFound, vswrDetected);
        }
        public bool ParseRETLOSS(string input, DataGridView grid, LogHandler logger)
        {
            var lines = input.Split("\r\n");
            bool result = true;
            foreach (var line in lines)
            {
                if (line.Contains("Ret Loss|"))
                {
                    var values = line.Split(new String[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 1; i < values.Length; i++)
                    {
                        bool parseResult = double.TryParse(values[i].Replace("dB", ""), out double value);
                        if (parseResult == true && value > 14)
                        {
                            result = true;

                        }
                        else
                        {
                            logger.tfailed.Add(new TestFailed { TestName = "Ret Loss : Path " + i, Value = values[i], Result = "FAIL", ErrorCodes = "TT045" });
                            result = false;
                        }
                        setGridValue(grid, DataRow.RET_Loss, i, values[i].Trim(), result);
                    }
                }
            }
            return result;
        }
        private bool ParseSFPReadings(string input, RichTextBox testLog, string Model, LogHandler logger)
        {
            bool sfpResult;
            bool sfpvalueisgood = true;
            int paths = 2;
            var lines = input.Split("\r\n");
            try
            {
                foreach (var line in lines)
                {
                    if (line.Contains("RxPow") || line.Contains("TxPow"))
                    {
                        var values = line.Split(new String[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 1; i <= paths; i++)
                        {
                            sfpResult = double.TryParse(values[i], out double value);
                            if (sfpResult)
                            {
                                if (value < -20)
                                {

                                }
                                if (value < -3 || value > 4)
                                {
                                    logger.tfailed.Add(new TestFailed { TestName = "SFP Test : Path " + i, Value = values, Result = "FAIL", ErrorCodes = "TT045" });
                                    this.Invoke(new MethodInvoker(delegate {
                                        testLog.SelectionColor = Color.Red;
                                        testLog.AppendText(values[0].Trim() + " on path " + i + " has failed: " + value + "\n");
                                        testLog.SelectionColor = Color.White;
                                    }));
                                    sfpvalueisgood = false;
                                }
                                else
                                {
                                    this.Invoke(new MethodInvoker(delegate {
                                        testLog.SelectionColor = Color.White;
                                        testLog.AppendText(values[0].Trim() + " on path " + i + " has passed: " + value + "\n");
                                    }));
                                }
                            }
                            else
                            {
                                this.Invoke(new MethodInvoker(delegate {
                                    testLog.SelectionColor = Color.DarkOrange;
                                    testLog.AppendText(values[0].Trim() + " on path " + i + " has returned: " + values[i].Trim() + "\n");
                                    testLog.SelectionColor = Color.White;
                                }));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                testLog.SelectionColor = Color.Yellow;
                testLog.AppendText("Error parsing boardOptShow\nCheck logs after testing\n");
                testLog.SelectionColor = Color.White;
            }

            return sfpvalueisgood;
        }
        private static string SSH_IP(string input, string RUID)
        {
            string sshIP = "-";
            /*var lines = input.Split("\r\n");
            foreach (var line in lines) {
                string[] values = line.Split(new String[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                if (line.Contains(RUID)) {
                    try {
                        sshIP = values[6];
                    }
                    catch(IndexOutOfRangeException ex) {
                        File.AppendAllText(@"C:\James's Projects\error.txt", ex.ToString() + "\n\n" + line);
                    }
                }
            }*/
            //mplane-info mplane-interfaces mplane-ipv6 REDACTED_IPV6
            if (input.Contains("No entries found"))
            {
                return sshIP;
            }
            var lines = input.Split("\r\n");
            var values = lines[1].Split(' ');
            sshIP = values[values.Length - 1];
            return sshIP;
        }
        private static string GetCSVAddress(string input)
        {
            string csvAddress = string.Empty;
            string[] lines = input.Split("\r\n");
            bool grab = false;
            foreach (string line in lines)
            {
                string[] values = line.Split(new String[] { " ", "/" }, StringSplitOptions.RemoveEmptyEntries);

                if (values.Contains("inet6") && grab == true && values.Contains("Scope:Link"))
                {
                    csvAddress = values[2];
                    grab = false;
                }
                if (values.Length == 5 && values[0] == "fh_0_0_0")
                {
                    grab = true;
                }
            }
            return csvAddress;
        }

        #endregion Log and Parsing Methods

        #region CSV Logging
        private bool RecentCsvExists(string serialNumber, string model)
        {
            if (!directory.ContainsKey(model))
                return false;

            string[] files = Directory.GetFiles(
                 directory[model] + @"CSV\",
                $"{serialNumber}_*.csv");

            DateTime cutoff =
                DateTime.Today.AddDays(-14);

            foreach (string file in files)
            {
                string name =
                    Path.GetFileNameWithoutExtension(file);

                string[] parts = name.Split('_');

                if (parts.Length != 2)
                    continue;

                if (DateTime.TryParseExact(
                        parts[1],
                        "yyyyMMdd",
                        null,
                        System.Globalization.DateTimeStyles.None,
                        out DateTime csvDate))
                {
                    if (csvDate >= cutoff)
                        return true;
                }
            }

            return false;
        }
        private void GetCSVLog(string serialNumber, string DU_IP, string sshIP, string csvIP, string logfile, string Model, RichTextBox testLog, int months, bool vswrDetected)
        {
            string[] config = File.ReadAllLines(@"GUI_Config.txt");
            string localUser = "";
            string localPass = "";
            string localIP = "";
            string serialCSVName = serialNumber + "_" + DateTime.Now.ToString("yyyyMMdd") + ".csv";
            foreach (string line in config)
            {
                string[] values = line.Split("=");
                switch (values[0])
                {
                    case "UserName":
                        localUser = values[1]; break;
                    case "Password":
                        localPass = values[1]; break;
                    case "IPAddress":
                        localIP = values[1]; break;
                }
            }
            Dictionary<string, int> failshutdown = new Dictionary<string, int>()
            {
                 { "0", 0 },
                { "1", 0 },
                { "2", 0 },
                { "3", 0 },
                { "4", 0 },
                { "5", 0 },
                { "6", 0 },
                { "7", 0 }

            };
            Dictionary<string, int> faillowgain = new Dictionary<string, int>()
           {
                 { "0", 0 },
                { "1", 0 },
                { "2", 0 },
                { "3", 0 },
                { "4", 0 },
                { "5", 0 },
                { "6", 0 },
                { "7", 0 }

            };
            Dictionary<string, int> failVSWR = new Dictionary<string, int>()
            {
                 { "0", 0 },
                { "1", 0 },
                { "2", 0 },
                { "3", 0 },
                { "4", 0 },
                { "5", 0 },
                { "6", 0 },
                { "7", 0 }

            };
            Dictionary<string, int> failclock = new Dictionary<string, int>()
            {
                 { "0", 0 },
                { "1", 0 },
                { "2", 0 },
                { "3", 0 },
                { "4", 0 },
                { "5", 0 },
                { "6", 0 },
                { "7", 0 }

            };
            Dictionary<string, int> failupdown = new Dictionary<string, int>()
            {
                 { "0", 0 },
                { "1", 0 },
                { "2", 0 },
                { "3", 0 },
                { "4", 0 },
                { "5", 0 },
                { "6", 0 },
                { "7", 0 }

            };
            Dictionary<string, int> failSOC = new Dictionary<string, int>()
{
     { "0", 0 },
    { "1", 0 },
    { "2", 0 },
    { "3", 0 },
    { "4", 0 },
    { "5", 0 },
    { "6", 0 },
    { "7", 0 }

};
            string[] pathArray = new string[8] { "0", "1", "2", "3", "4", "5", "6", "7" };
            DateTime month = DateTime.Now.AddMonths(months * -1);
            DateTime days = DateTime.Now.AddDays(-2);
            DateTime dateTime;
            DateTime timeCheck;
            int count = 0;
            List<string> alarmDetected = new List<string>();

            var client = new SshClient(DU_IP, "REDACTED_USER", "REDACTED_PASSWORD");
            this.Invoke(new MethodInvoker(delegate {
                testLog.AppendText("CSV Logging has begun\n");

            }));
            client.Connect();
            var stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
            string reader = stream.Read();

            timeCheck = DateTime.Now.AddMinutes(2);
            stream.WriteLine("./conn rmp");
            Thread.Sleep(500);
            while (!reader.Contains("sapp@uadpf") && DateTime.Now < timeCheck)
            {
                Thread.Sleep(10000);
                reader += stream.Read();
            }
            File.AppendAllText(logfile, reader);

            stream.WriteLine("rm -rf /home/sapp/.ssh/known_hosts");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);


            stream.WriteLine("ssh user@" + sshIP);
            Thread.Sleep(5000);
            reader = stream.Read();
            File.AppendAllText(logfile, reader); int sshCountDown = 3;
            /*if(reader.Contains("No route to host"))
            {
                this.Invoke(new MethodInvoker(delegate
                {
                    testLog.AppendText("No route to host. Ending test");
                }));
                goto EndTest;
            }*/
            while (sshCountDown > 0 && !reader.Contains("fingerprint"))
            {
                sshCountDown--;

                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Failed to SSH into unit. Number of retries left " + sshCountDown.ToString() + "\n");
                }));
                Thread.Sleep(60000);
                stream.WriteLine("ssh user@" + sshIP);
                Thread.Sleep(5000);
                reader = stream.Read();
            }
            if (sshCountDown == 0)
            {
                this.Invoke(new MethodInvoker(delegate {
                    testLog.SelectionColor = Color.Red;
                    testLog.AppendText("Failed to SSH into unit. Failing test\n");
                    testLog.SelectionColor = Color.White;
                }));
                goto StopCSV;
            }
            this.Invoke(new MethodInvoker(delegate {
                testLog.AppendText("Successful SSH into unit\n");
            }));
            stream.WriteLine("yes");
            Thread.Sleep(1000);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);
            if (!reader.Contains("computer system and network"))
            {
                client.Disconnect();
                this.Invoke(new MethodInvoker(delegate {
                    testLog.SelectionColor = Color.Red;
                    testLog.AppendText("Infinite y of death. Please wait 3 minutes or until 3 or 4 LEDs are blinking\n");
                    testLog.SelectionColor = Color.White;
                }));
                goto StopCSV;
            }

            stream.WriteLine("REDACTED_PASSWORD");
            Thread.Sleep(12000);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            int permissionCountDown2 = 3;
            while (reader.Contains("Permission denied, please try again.") && permissionCountDown2 > 0)
            {
                permissionCountDown2--;
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Permission denied. Number of retries left: " + permissionCountDown2 + "\n");
                }));
                stream.WriteLine("");
                Thread.Sleep(12000);
                stream.WriteLine("");
                Thread.Sleep(1000);
                stream.WriteLine("ssh user@" + sshIP);
                Thread.Sleep(5000);
                reader = stream.Read();
                stream.WriteLine("REDACTED_PASSWORD");
                Thread.Sleep(12000);
                reader = stream.Read();
                File.AppendAllText(logfile, reader);
            }
            if (permissionCountDown2 == 0)
            {
                this.Invoke(new MethodInvoker(delegate {
                    testLog.SelectionColor = Color.Red;
                    testLog.AppendText("Permission denied. Ending test\n");
                    testLog.SelectionColor = Color.White;
                }));
                goto StopCSV;
            }

            stream.WriteLine("su -");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            stream.WriteLine("REDACTED_PASSWORD");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            stream.WriteLine("mortem mortem");
            Thread.Sleep(5000);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            stream.WriteLine("ushell");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);


            stream.WriteLine("Alarm_HisGet 2");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            DateTime timeToCheck = DateTime.Now.AddMinutes(2);
            while (!reader.Contains("Alarm History has been saved") && DateTime.Now < timeToCheck)
            {
                Thread.Sleep(10000);
                stream.WriteLine("");
                reader = stream.Read();
            }
            if (timeToCheck <= DateTime.Now)
            {
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Failed to get CSV Log. Ending logging\n");

                }));
                goto StopCSV;
            }

            stream.WriteLine("exit");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            stream.WriteLine("cd /tmp");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            stream.WriteLine("chmod 777 rulog.csv");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            stream.WriteLine("mv rulog.csv " + serialCSVName);
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            stream.WriteLine("exit");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            stream.WriteLine("exit");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            stream.WriteLine("exit");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            stream.WriteLine(@"cd CSV");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            stream.WriteLine("rm -rf /home/sysadmin/.ssh/known_hosts");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            stream.WriteLine("sftp user@[" + csvIP + "%ens1f0]");
            Thread.Sleep(5000);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            if (reader.Contains("No route to host"))
            {
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("No route to host. Ending CSV logging\n");

                }));
                goto StopCSV;
            }

            stream.WriteLine("yes");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            if (!reader.Contains("computer system and network"))
            {
                client.Disconnect();
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Infinite y of death. Please wait 3 minutes or until 3 or 4 LEDs are blinking\n");

                }));
                goto StopCSV;
            }

            stream.WriteLine("REDACTED_PASSWORD");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            stream.WriteLine("cd /tmp");
            Thread.Sleep(500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            stream.WriteLine("get " + serialCSVName);
            Thread.Sleep(1500);
            reader = stream.Read();
            File.AppendAllText(logfile, reader);

            int countDown = 10;
            string csvAddress = @"/var/home/REDACTED_USER/";

            while (!reader.Contains("100%") && countDown > 0)
            {
                countDown--;
                stream.WriteLine("");
                Thread.Sleep(5000);
                reader = stream.Read();
                File.AppendAllText(logfile, reader);
            }
            var sftpClient = new SftpClient(DU_IP, "REDACTED_USER", "REDACTED_PASSWORD");
            string moveTo = @"C:\Log\CSV\";
            string localAddress = moveTo + serialCSVName;
            try
            {
                sftpClient.Connect();
                Stream fileStream = File.Create(localAddress);
                sftpClient.DownloadFile(csvAddress + serialCSVName, fileStream);
                fileStream.Dispose();
                sftpClient.Delete(csvAddress + serialCSVName);
                sftpClient.Disconnect();
            }
            catch (Exception ex)
            {
                //File.AppendAllText(ex.ToString() + "\n", logfile);
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Exception thrown while attempting to down CSV file.\nContact Engineering.");
                }));
                Task.Run(() => { MessageBox.Show("Exception thrown while attempting to down CSV file.\nContact Engineering.\n" + ex.ToString()); });
            }

            using (TextFieldParser parser = new TextFieldParser(localAddress))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                while (!parser.EndOfData)
                {
                    //Process row
                    string[] fields = parser.ReadFields();
                    foreach (string field in fields)
                    {
                        //TODO: Process field
                        if (field.Contains("ShutDown") && fields[6].Contains("OCCUR"))
                        {
                            string[] pathValues = fields[8].Split("(");
                            var values = fields[3].Split(" ");
                            dateTime = DateTime.Parse(values[0]);
                            if (dateTime < days && month < dateTime)
                            {
                                failshutdown[pathValues[0].Trim()]++;
                            }

                        }
                        else if (field.Contains("VswrFail(MJ)") && fields[6].Contains("OCCUR") && vswrDetected == true)
                        {
                            var values = fields[3].Split(" ");
                            string[] pathValues = fields[8].Split("(");
                            dateTime = DateTime.Parse(values[0]);
                            if (dateTime < days && month < dateTime)
                            {
                                failVSWR[pathValues[0].Trim()]++;
                            }

                        }
                        else if (field.Contains("ClockFail") && fields[6].Contains("OCCUR"))
                        {
                            var values = fields[3].Split(" ");
                            string[] pathValues = fields[8].Split("(");
                            dateTime = DateTime.Parse(values[0]);
                            if (dateTime < days && month < dateTime)
                            {
                                failclock[pathValues[0].Trim()]++;
                            }

                        }
                        else if (field.Contains("LowGain") && !field.Contains("LowGainSymptom") && fields[6].Contains("OCCUR"))
                        {
                            var values = fields[3].Split(" ");
                            string[] pathValues = fields[8].Split("(");
                            dateTime = DateTime.Parse(values[0]);
                            if (dateTime < days && month < dateTime)
                            {
                                faillowgain[pathValues[0].Trim()]++;
                            }

                        }
                        else if (field.Contains("UpDownConvError") && fields[6].Contains("OCCUR"))
                        {
                            var values = fields[3].Split(" ");
                            string[] pathValues = fields[8].Split("(");
                            dateTime = DateTime.Parse(values[0]);
                            if (dateTime < days && month < dateTime)
                            {
                                failupdown[pathValues[0].Trim()]++;
                            }
                        }
                        else if (field.Contains("SOCFail") && fields[6].Contains("OCCUR"))
                        {
                            var values = fields[3].Split(" ");
                            string[] pathValues = fields[8].Split("(");
                            dateTime = DateTime.Parse(values[0]);
                            if (dateTime < days && month < dateTime) { failSOC[pathValues[0].Trim()]++; }
                        }
                    }
                }
            }
            foreach (string path in pathArray)
            {
                if (failshutdown[path] > 0)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        //this part below applies to CarrierA only    |---------------------| 
                        testLog.AppendText("PATH " + (Int32.Parse(path) + 1) + " ShutDown Alarm Occur " + failshutdown[path] + " times in the past " + months + " Months\n");
                        //testLog.AppendText("Path " + path + " has failed " + failCount[path] + " times in the past " + months + " months\n"); //Uncomment this for CarrierB ORAN

                    }));
                    //Console.WriteLine("PATH " + (Int32.Parse(path) + 1) + " ShutDown Alarm Occur " + failshutdown[path] + " times in the past " + months + " Months");
                }
                if (faillowgain[path] > 0)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        //this part below applies to CarrierA only    |---------------------| 
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("PATH: " + (Int32.Parse(path) + 1) + " Low Gain Alarm Occur: " + faillowgain[path] + " Times Past " + months + " Months\n");
                        testLog.SelectionColor = Color.White;
                        //testLog.AppendText("Path " + path + " has failed " + failCount[path] + " times in the past " + months + " months\n"); //Uncomment this for CarrierB ORAN

                    }));
                    //Console.WriteLine("PATH: " + (Int32.Parse(path) + 1) + " Low Gain Alarm Occur: " + faillowgain[path] + " Times Past " + months + " Months");
                }
                if (failVSWR[path] > 0)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        //this part below applies to CarrierA only    |---------------------| 
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("PATH: " + (Int32.Parse(path) + 1) + " VSWR Alarm Occur: " + failVSWR[path] + " Times Past " + months + " Months\n");
                        testLog.SelectionColor = Color.White;
                        //testLog.AppendText("Path " + path + " has failed " + failCount[path] + " times in the past " + months + " months\n"); //Uncomment this for CarrierB ORAN

                    }));
                    //Console.WriteLine("PATH: " + (Int32.Parse(path) + 1) + " VSWR Alarm Occur: " + failVSWR[path] + " Times Past " + months + " Months");
                }
                if (failclock[path] > 0)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        //this part below applies to CarrierA only    |---------------------| 
                        testLog.AppendText("PATH: " + (Int32.Parse(path) + 1) + " Clock Fail Alarm Occur: " + failclock[path] + " Times Past " + months + " Months\n");
                        testLog.SelectionColor = Color.White;
                        //testLog.AppendText("Path " + path + " has failed " + failCount[path] + " times in the past " + months + " months\n"); //Uncomment this for CarrierB ORAN

                    }));
                    //Console.WriteLine("PATH: " + (Int32.Parse(path) + 1) + " Clock Fail Alarm Occur: " + failclock[path] + " Times Past " + months + " Months");
                }
                if (failupdown[path] > 0)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        //this part below applies to CarrierA only    |---------------------| 
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("UPDOWN: " + (Int32.Parse(path) + 1) + " UpDownConvError Alarm Occur: " + failupdown[path] + " Times Past " + months + " Months\n");
                        testLog.SelectionColor = Color.White;
                        //testLog.AppendText("Path " + path + " has failed " + failCount[path] + " times in the past " + months + " months\n"); //Uncomment this for CarrierB ORAN

                    }));
                    //Console.WriteLine("PATH: " + (Int32.Parse(path) + 1) + " Clock Fail Alarm Occur: " + failclock[path] + " Times Past " + months + " Months");
                }
                if (failSOC[path] > 0)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        //this part below applies to CarrierA only    |---------------------| 
                        testLog.SelectionColor = Color.Red;
                        //testLog.AppendText("Path " + (Int32.Parse(path) + 1) + " has failed " + failshutdown[path] + " times in the past " + months + " months\n");
                        testLog.AppendText("PATH " + path + " SOCFail Alarm Occur " + failSOC[path] + " times in the past " + months + " Months\n"); //Uncomment this for CarrierB ORAN
                        testLog.SelectionColor = Color.White;
                    }));
                    //Console.WriteLine("PATH " + (Int32.Parse(path) + 1) + " ShutDown Alarm Occur " + failshutdown[path] + " times in the past " + months + " Months");
                }
            }
            if (!Directory.Exists(directory[Model]))
            {
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("T Drive unavailable\n");
                }));
            }
            else
            {
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("T Drive is available. Adding CSV to T drive\n");
                }));
                moveTo = directory[Model] + @"CSV\";
                if (File.Exists(moveTo + serialCSVName))
                {
                    File.Delete(moveTo + serialCSVName);
                    File.Move(localAddress, moveTo + serialCSVName);
                }
                else
                {
                    File.Move(localAddress, moveTo + serialCSVName);
                }
            }

        StopCSV:;
        }

        #endregion CSV Logging
        // In-progress refactor target: not called anywhere yet, but intended to replace the
        // repeated stream.WriteLine(...) / Thread.Sleep(1000) / stream.Read() pattern used
        // throughout this file (40+ call sites) with a condition-based wait — checking for an
        // expected endpoint string and live connection instead of a fixed delay.
        #region SSH_Command
        private string SSH_Command(ShellStream stream, SshClient client, string command, string endpoint, int timeout = 15)
        {
            string result = string.Empty;
            try
            {
                stream.Read();
                stream.WriteLine(command);
                while (!result.Contains(endpoint) && timeout > 0 && client.IsConnected)
                {
                    result += stream.Read();
                }
            }
            catch (SshConnectionException)
            {
                MessageBox.Show("client connection lost. Restart testing when 3 or 4 LEDs are blinking");
            }
            catch (Exception ex)
            {
                result = ex.ToString();
            }

            return result;
        }
        #endregion SSH_Command


        //----------------------Main Test---------------------
        public void StartTest(string DU_IP, string RUID, DataGridView grid, RichTextBox testLog, TextBox scannedSN, string Model, System.Timers.Timer timer, Label timeStamp)
        {
            //Stuff and things go here
            string serialNumber = scannedSN.Text;
            string postBooter = "";
            string sshIP = "";
            bool txPowerPassed = true;
            bool returnLossPassed = true;
            bool RetTestisPassed = true;
            bool AlarmpnotPresent = true;
            bool IsFirmwareValid = true;
            bool IsSshConnected = true;
            bool sfpPresent = true;
            bool vswrDetected = false;
            string firmwarePath = @"C:\Carrier Oran Settings\Firmware.txt";
            string expectedFirmware = "";
            string Firmwareactive = "";
            LogHandler logger = new LogHandler();
            // Ensure the directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(firmwarePath));

            // Check if file exists, create if it doesn't
            if (!File.Exists(firmwarePath))
            {
                File.WriteAllText(firmwarePath, ""); // create empty file
            }
            if (Model == "PCS") { postBooter = "postbooter.a.rf_model_a.0"; } else if (Model == "LOLO") { postBooter = "postbooter.a.rf_model_b.0"; }
            DateTime timeCheck;

            #region MainTest
            var client = new SshClient(DU_IP, "REDACTED_USER", "REDACTED_PASSWORD");
            client.ConnectionInfo.Timeout = TimeSpan.FromSeconds(600); //This doesn't help
            {
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Test was started at " + DateTime.Now.ToString("hh:mm tt") + "\nPlug in the power now\n");
                }));

                #region AssignSerialNumber
                List<string> LogFileList = LogBuilder(scannedSN.Text, Model);
                string logfile = LogFileList[0];
                File.AppendAllText(logfile, "\nLog Started\r**********************************************************\r");
                client.ConnectionInfo.Timeout = TimeSpan.FromSeconds(600);
                try
                {
                    client.Connect();
                }
                catch (SocketException e)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Unable to connect to vDU\n");
                    }));
                    IsSshConnected = false;
                    logger.tfailed.Add(new TestFailed { TestName = "Connection: ", Value = "NA", Result = "FAIL", ErrorCodes = "TT023" });
                    goto EndTest;
                }
                var stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                Thread.Sleep(5000);
                string reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                timeCheck = DateTime.Now.AddMinutes(2);
                stream.WriteLine("./conn dmp");
                Thread.Sleep(500);
                while (!reader.Contains("sapp@uadpf") && DateTime.Now < timeCheck)
                {
                    Thread.Sleep(10000);
                    reader += stream.Read();
                }
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                stream.WriteLine("nrconfd_cli -u vsmuser");
                Thread.Sleep(1000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                stream.WriteLine("set paginate false");
                Thread.Sleep(1000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("show table managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info serial-number");
                Thread.Sleep(1000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                string[] lines = reader.Split("\r\n");

                if (reader.Contains(scannedSN.Text))
                {

                    foreach (string line in lines)
                    {
                        if (line.Contains(scannedSN.Text))
                        {
                            string[] values = line.Split(new String[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                            RUID = values[0];
                            this.Invoke(new MethodInvoker(delegate {
                                testLog.AppendText("Serial number already assigned to " + RUID + "\n");
                            }));
                            goto SkipAssignSerial;
                        }
                    }
                }
                else
                {
                    stream.WriteLine("configure");
                    Thread.Sleep(1000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                    stream.WriteLine("set managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info " + RUID + " serial-number " + scannedSN.Text);
                    Thread.Sleep(1000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                    stream.WriteLine("commit");
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                    stream.WriteLine("exit");
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                    stream.WriteLine("show table managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info serial-number");
                    Thread.Sleep(1000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                }
            #endregion AssignSerialNumber

            SkipAssignSerial:;
                timer.Stop();
                //var pwrresult = MessageBox.Show("Is the power turning on for the unit?\n" + Model, "",
                //       MessageBoxButtons.YesNo,
                //       MessageBoxIcon.Question);
                //if (pwrresult == DialogResult.No) {
                //    this.Invoke(new MethodInvoker(delegate {
                //        testLog.AppendText("Ending Test\n");
                //    }));
                //    goto EndTest;
                //}
                DialogResult pwrresult = DialogResult.None;

                this.Invoke(() =>
                {
                    pwrresult = MessageBox.Show(
                        this,
                        $"Is the power turning on for the unit?\n{Model}",
                        "",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                });

                // Background thread is blocked here until the Invoke completes,
                // and the Invoke won't return until the user clicks Yes or No.

                if (pwrresult == DialogResult.No)
                {
                    this.Invoke(() =>
                    {
                        testLog.AppendText("Ending Test\n");
                    });

                    goto EndTest;
                }
                if (!client.IsConnected)
                {
                    client.Connect();
                    stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                    timeCheck = DateTime.Now.AddMinutes(2);
                    stream.WriteLine("./conn dmp");
                    Thread.Sleep(500);
                    while (!reader.Contains("sapp@uadpf") && DateTime.Now < timeCheck)
                    {
                        Thread.Sleep(10000);
                        reader += stream.Read();
                    }
                    stream.WriteLine("nrconfd_cli -u vsmuser");
                    Thread.Sleep(1000);
                    stream.WriteLine("set paginate false");
                    Thread.Sleep(1000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                }
                timer.Start();
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Moving to pick up IP address\n");
                }));

                #region PickupIPAddress
                sshIP = "-";
                DateTime timeToCheck = DateTime.Now.AddMinutes(12);
                stream.WriteLine("");
                Thread.Sleep(1000);
                stream.Read();
                reader = string.Empty;
                while (timeToCheck > DateTime.Now && sshIP == "-")
                {
                    //show managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info {RUID} mplane-info mplane-interfaces mplane-ipv6  S518627514
                    stream.Write($"show managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info {RUID} mplane-info mplane-interfaces mplane-ipv6\r");
                    Thread.Sleep(1000);
                    reader = stream.Read();
                    sshIP = SSH_IP(reader, RUID);
                    if (sshIP != "-")
                    {
                        break;
                    }
                    Thread.Sleep(60000);

                }
                if (sshIP == "-")
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("Unit failed to connect to VDU. Ending test\n");
                        testLog.SelectionColor = Color.White;
                    }));
                    IsSshConnected = false;
                    logger.tfailed.Add(new TestFailed { TestName = "Connection: ", Value = "NA", Result = "FAIL", ErrorCodes = "TT023" });
                    goto EndTest;
                }

                #endregion PickupIPAddress
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("IP Address picked up\n");
                }));
                Thread.Sleep(2000);
                stream.WriteLine("exit");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("exit");
                Thread.Sleep(500);
                reader = stream.Read();
                #region ConnectToRoot
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Connecting to root\n");
                }));
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                timeCheck = DateTime.Now.AddMinutes(2);
                stream.WriteLine("./conn rmp");
                Thread.Sleep(500);
                while (!reader.Contains("sapp@uadpf") && DateTime.Now < timeCheck)
                {
                    Thread.Sleep(10000);
                    reader += stream.Read();
                }
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("rm -rf /home/sapp/.ssh/known_hosts");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");


                stream.WriteLine("ssh user@" + sshIP);
                Thread.Sleep(5000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                int sshCountDown = 4;
                if (reader.Contains("WARNING: REMOTE HOST IDENTIFICATION HAS CHANGED!"))
                {

                    stream.WriteLine("rm -rf /home/sapp/.ssh/known_hosts");
                    Thread.Sleep(500);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");


                    stream.WriteLine("ssh user@" + sshIP);
                    Thread.Sleep(5000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                }
                while (sshCountDown > 0 && reader.Contains("No route to host"))
                {
                    sshCountDown--;

                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Failed to SSH into unit. Number of retries left " + sshCountDown.ToString() + "\n");
                    }));
                    Thread.Sleep(60000);
                    stream.WriteLine("rm -rf /home/sapp/.ssh/known_hosts");
                    Thread.Sleep(500);
                    stream.WriteLine("ssh user@" + sshIP);
                    Thread.Sleep(5000);
                    reader = stream.Read();
                }
                if (sshCountDown == 0)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("Failed to SSH into unit. Failing test\n");
                        testLog.SelectionColor = Color.White;
                    }));
                    goto EndTest;
                }
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Successful SSH into unit\n");
                }));
                stream.WriteLine("yes");
                Thread.Sleep(1000);
                reader = stream.Read();
                if (!reader.Contains("computer system and network"))
                {
                    client.Disconnect();
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("Infinite y of death. Please wait 3 minutes or until 3 or 4 LEDs are blinking\n");
                        testLog.SelectionColor = Color.White;
                    }));
                    goto EndTest;
                }
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");


                stream.WriteLine("REDACTED_PASSWORD");
                Thread.Sleep(12000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                int permissionCountDown2 = 3;
                while (reader.Contains("Permission denied, please try again.") && permissionCountDown2 > 0)
                {
                    permissionCountDown2--;
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Permission denied. Number of retries left: " + permissionCountDown2 + "\n");
                    }));
                    stream.WriteLine("");
                    Thread.Sleep(12000);
                    stream.WriteLine("");
                    Thread.Sleep(1000);
                    stream.WriteLine("ssh user@" + sshIP);
                    Thread.Sleep(5000);
                    reader = stream.Read();
                    stream.WriteLine("REDACTED_PASSWORD");
                    Thread.Sleep(12000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                }
                if (permissionCountDown2 == 0)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("Permission denied. Ending test\n");
                        testLog.SelectionColor = Color.White;
                    }));
                    goto EndTest;
                }

                stream.WriteLine("su -");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("REDACTED_PASSWORD");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("setenv p BOOT_CONSOLE_LOG YES");
                Thread.Sleep(3000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("getinv");
                Thread.Sleep(3000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                lines = reader.Split("\r\n");
                foreach (string line in lines)
                {
                    if (line.Contains("HW SN"))
                    {
                        string[] values = line.Split(new String[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.AppendText("Internal serial number is " + values[3] + "\n");
                        }));
                        if (scannedSN.Text == values[3])
                        {
                            switch (Model)
                            {
                                case "PCS":
                                    this.Invoke(new MethodInvoker(delegate {
                                        serialNumber_PCS.HeaderText = values[3];
                                    }));
                                    break;
                                case "LOLO":
                                    this.Invoke(new MethodInvoker(delegate {
                                        serialNumber_LOLO.HeaderText = values[3];
                                    }));
                                    break;
                            }
                        }
                        else
                        {
                            timer.Stop();
                            var snResult = MessageBox.Show("Scanned serial and internal serial number doesn't match. Continue?", "",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);
                            if (snResult == DialogResult.No)
                            {
                                this.Invoke(new MethodInvoker(delegate {
                                    testLog.AppendText("Ending Test\n");
                                }));
                                goto EndTest;
                            }

                            timer.Start();
                        }

                    }
                }


                if (!client.IsConnected)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Connection lost. Attempting to log into root");
                    }));
                    client.Connect();
                    stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                    Thread.Sleep(2000);
                    timeCheck = DateTime.Now.AddMinutes(2);
                    stream.WriteLine("./conn rmp");
                    Thread.Sleep(500);
                    while (!reader.Contains("sapp@uadpf") && DateTime.Now < timeCheck)
                    {
                        Thread.Sleep(10000);
                        reader += stream.Read();
                    }
                    stream.WriteLine("rm -rf /home/sapp/.ssh/known_hosts");
                    Thread.Sleep(1000);
                    stream.WriteLine("ssh user@" + sshIP);
                    Thread.Sleep(5000);
                    reader = stream.Read();
                    if (reader.Contains("No route to host"))
                    {
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.SelectionColor = Color.Red;
                            testLog.AppendText("Unable to SSH into unit. Failing test\n");
                            testLog.SelectionColor = Color.White;

                        }));
                        IsSshConnected = false;
                        logger.tfailed.Add(new TestFailed { TestName = "Connection: ", Value = "NA", Result = "FAIL", ErrorCodes = "TT023" });
                        goto EndTest;
                    }
                    stream.WriteLine("yes");
                    Thread.Sleep(1000);
                    if (!reader.Contains("computer system and network"))
                    {
                        client.Disconnect();
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.AppendText("Infinite y of death. Please wait 3 minutes or until 3 or 4 LEDs are blinking\n");
                        }));
                        goto EndTest;
                    }
                    Thread.Sleep(200);
                    stream.WriteLine("REDACTED_PASSWORD");
                    Thread.Sleep(12000);
                    reader = stream.Read();
                    if (reader.Contains("Permission denied"))
                    {
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.SelectionColor = Color.Red;
                            testLog.AppendText("Permission denied. Ending test\n");
                            testLog.SelectionColor = Color.White;
                        }));
                        goto EndTest;
                    }
                    stream.WriteLine("su -");
                    Thread.Sleep(1000);
                    stream.WriteLine("REDACTED_PASSWORD");
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                }
                if (reader.Contains("SAFE"))
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("Unit in SAFE mode. Ending test\n");
                        testLog.SelectionColor = Color.White;
                    }));
                    goto EndTest;
                }
                reader = string.Empty;
                int countdown = 5;
                stream.WriteLine("disfilever /mnt/storage/slot_1/pkg/bin/*");
                while (!reader.Contains("root@") && countdown > 0)
                {
                    reader += stream.Read();
                    Thread.Sleep(5000);
                }
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                reader = string.Empty;
                countdown = 5;

                stream.WriteLine("disfilever /mnt/storage/slot_2/pkg/bin/*");
                while (!reader.Contains("root@") && countdown > 0)
                {
                    reader += stream.Read();
                    Thread.Sleep(5000);
                }
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                string[] forbiddenKeywords = { "factory", "factory mode", "facotry mode", "safe mode", "Safe Mode" };
                reader = string.Empty;
                stream.Read();
                stream.Write("gettail 0\r");
                timeCheck = DateTime.Now.AddSeconds(30);
                while (!reader.Contains("root@") && DateTime.Now < timeCheck)
                {
                    reader += stream.Read();
                }
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                try
                {
                    if (Model == "PCS" || Model == "LOLO")
                    {
                        foreach (string line in reader.Split("\r\n"))
                        {
                            if (line.Contains("postbooter"))
                            {
                                var values = line.Split(new String[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                                if (values[1] != postBooter)
                                {
                                    this.Invoke(new MethodInvoker(delegate {
                                        testLog.SelectionColor = Color.Red;
                                        testLog.AppendText("Postbooter not installed\rSend to repair once test is finished\n");
                                        testLog.SelectionColor = Color.White;
                                    }));
                                }
                            }
                            if (line.TrimStart().StartsWith("lptmc", StringComparison.OrdinalIgnoreCase))
                            {
                                var parts = line.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                if (parts.Length > 1)
                                {
                                    string blockName = parts[1].ToLower();

                                    foreach (string keyword in forbiddenKeywords)
                                    {
                                        if (blockName.Contains(keyword.ToLower()))
                                        {
                                            this.Invoke(new MethodInvoker(delegate {
                                                testLog.SelectionColor = Color.Red;
                                                testLog.AppendText("LPTMC�module Contains Factory Mode..!\n");
                                                testLog.SelectionColor = Color.White;
                                            }));
                                            break;
                                        }
                                        else
                                        {
                                            this.Invoke(new MethodInvoker(delegate {
                                                testLog.SelectionColor = Color.LightGreen;
                                                testLog.AppendText("LPTMC Mode is : " + parts[1].ToLower() + "�\n");
                                                testLog.SelectionColor = Color.White;
                                            }));
                                        }
                                    }
                                }

                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    File.AppendAllText(logfile, ex.ToString() + Environment.NewLine);
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.Orange;
                        testLog.AppendText("Exception Thrown for gettail 0. Continuing test...\n");
                        testLog.SelectionColor = Color.White;
                    }));
                }


                stream.WriteLine("printenv");
                Thread.Sleep(1000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                countdown = 3;
                while (!reader.Contains("BOOT_CONSOLE_LOG=YES") && countdown > 0)
                {
                    stream.WriteLine("set BOOT_CONSOLE_LOG YES");
                    Thread.Sleep(3000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("printenv");
                    Thread.Sleep(2000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                    countdown--;
                }
                if (countdown == 0)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("Unit is not unlocked. Stopping test.\rTake over manually\n");
                        testLog.SelectionColor = Color.White;
                    }));
                    goto EndTest;
                }
                if (reader.Contains("slot#2=-0") || reader.Contains("slot#1=-0"))
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("Flash memory corrupted. Ending test\n");
                        testLog.SelectionColor = Color.White;
                    }));
                    goto EndTest;
                }
                else if (reader.Contains("slot#2=19") || reader.Contains("slot#1=19"))
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("Environment corruption. Ending test\n");
                        testLog.SelectionColor = Color.White;
                    }));
                    goto EndTest;
                }
                else if (reader.Contains("slot#1=10") || reader.Contains("slot#2=10"))
                {
                    ///need to write up code to wait for firmware to finish updating
                    ///if after everything and slot 2 still equals 01, end the test and have the user take over manually
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("The Firmware is updating\rLet the Unit Reboot\n");
                    }));
                    timeToCheck = DateTime.Now.AddMinutes(13);
                    while (reader.Contains("=10") && !reader.Contains("send disconnect: Broken pipe") && !reader.Contains("Ru Reset") && timeToCheck > DateTime.Now)
                    {
                        stream.WriteLine("printenv");
                        Thread.Sleep(5000);
                        reader = stream.Read();
                    }
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                    if (timeToCheck <= DateTime.Now)
                    {
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.SelectionColor = Color.Red;
                            testLog.AppendText("Pipe was not broken. Ending test\n");
                            testLog.SelectionColor = Color.White;
                        }));
                        goto EndTest;
                    }
                    else
                    {
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.AppendText("Pipe has been broken. Starting 5 minute wait\n" + DateTime.Now.ToString("hh:mm tt"));
                        }));
                    }
                    /* if (reader.Contains("send disconnect: Broken pipe")) {
                         this.Invoke(new MethodInvoker(delegate {
                             testLog.AppendText("Pipe has been broken. Starting 4 minute wait\n" + DateTime.Now.ToString("hh:mm tt"));
                         }));
                     } else {
                         this.Invoke(new MethodInvoker(delegate {
                             testLog.SelectionColor = Color.Red;
                             testLog.AppendText("Pipe was not broken. Ending test\n");
                             testLog.SelectionColor = Color.White;
                         }));
                         goto EndTest;
                     }*/
                    client.Disconnect();
                    reader = string.Empty;
                    Thread.Sleep(300000);

                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("\nChecking Firmware status...\n");
                    }));
                    if (client.IsConnected)
                    {
                        stream.WriteLine("exit");
                        Thread.Sleep(500);
                        reader = stream.Read();
                        File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                        Thread.Sleep(30000);

                    }
                    else
                    {
                        client.Connect();
                        stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                        Thread.Sleep(300);
                    }

                    timeCheck = DateTime.Now.AddMinutes(2);
                    stream.WriteLine("./conn rmp");
                    Thread.Sleep(500);
                    while (!reader.Contains("sapp@uadpf") && DateTime.Now < timeCheck)
                    {
                        Thread.Sleep(10000);
                        reader += stream.Read();
                    }
                    Thread.Sleep(500);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("rm -rf /home/sapp/.ssh/known_hosts");
                    Thread.Sleep(500);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("ssh user@" + sshIP);
                    Thread.Sleep(5000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    if (!reader.Contains("fingerprint"))
                    {
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.SelectionColor = Color.Red;
                            testLog.AppendText("Failed to SSH Into Unit. Ending test\n");
                            testLog.SelectionColor = Color.White;
                        }));
                        goto EndTest;
                    }

                    stream.WriteLine("yes");
                    Thread.Sleep(5000);
                    reader = stream.Read();

                    if (!reader.Contains("computer system and network"))
                    {
                        client.Disconnect();
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.SelectionColor = Color.Red;
                            testLog.AppendText("Infinite y of death. Please wait 3 minutes or until 3 or 4 LEDs are blinking\n");
                            testLog.SelectionColor = Color.White;
                        }));
                        goto EndTest;
                    }
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("REDACTED_PASSWORD");
                    Thread.Sleep(12000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    //This while loop below may not work based on some new things that have come up during troubleshooing for RU Aging
                    permissionCountDown2 = 2;
                    while (reader.Contains("Permission denied, please try again.") && permissionCountDown2 > 0)
                    {
                        permissionCountDown2--;
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.AppendText("Permission denied. Number of retries left: " + permissionCountDown2 + "\n");
                        }));
                        stream.WriteLine("");
                        Thread.Sleep(12000);
                        stream.WriteLine("");
                        Thread.Sleep(1000);
                        stream.WriteLine("ssh user@" + sshIP);
                        Thread.Sleep(5000);
                        reader = stream.Read();
                        if (!reader.Contains("fingerprint"))
                        {
                            this.Invoke(new MethodInvoker(delegate {
                                testLog.SelectionColor = Color.Red;
                                testLog.AppendText("Failed to SSH Into Unit. Ending test\n");
                                testLog.SelectionColor = Color.White;
                            }));
                            goto EndTest;
                        }
                        stream.WriteLine("REDACTED_PASSWORD");
                        Thread.Sleep(500);
                        reader = stream.Read();
                        File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                        Thread.Sleep(12000);
                    }
                    if (permissionCountDown2 == 0)
                    {
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.AppendText("Permission denied. Ending test\n");
                        }));
                        goto EndTest;
                    }
                    stream.WriteLine("su -");
                    Thread.Sleep(500);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("REDACTED_PASSWORD");
                    Thread.Sleep(500);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("printenv");
                    Thread.Sleep(500);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    if (reader.Contains("slot#2=01") || reader.Contains("slot#1=01"))
                    {
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.AppendText("Firmware has upgraded. Continuing test...\n");
                        }));
                    }
                }
                else if (reader.Contains("slot#2=01") || reader.Contains("slot#1=01"))
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Firmware is good. Continuing test...\n");
                    }));
                }
                string activeSlot = "";
                if (reader.Contains("slot#2=01"))
                {
                    activeSlot = "2";
                }
                else if (reader.Contains("slot#1=01"))
                {
                    activeSlot = "1";
                }


                stream.WriteLine("disfilever /mnt/storage/slot_1/pkg/bin/*");
                while (!reader.Contains("root@") && countdown > 0)
                {
                    reader += stream.Read();
                    Thread.Sleep(5000);
                }
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                reader = string.Empty;
                countdown = 5;

                stream.WriteLine("disfilever /mnt/storage/slot_2/pkg/bin/*");
                while (!reader.Contains("root@") && countdown > 0)
                {
                    reader += stream.Read();
                    Thread.Sleep(5000);
                }
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("gettail 0");
                Thread.Sleep(2000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                if (File.Exists(firmwarePath))
                {
                    expectedFirmware = File.ReadAllText(firmwarePath).Trim();
                }
                foreach (string line in reader.Split("\r\n"))
                {
                    if (line.Contains("booter." + activeSlot))
                    {
                        var values = line.Split(new String[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                        Firmwareactive = values[4];
                        if (!values[4].Contains(expectedFirmware))
                        {
                            this.Invoke(new MethodInvoker(delegate {
                                testLog.SelectionColor = Color.Red;
                                testLog.AppendText("Slot#" + activeSlot + " did not update firmware\n");
                                testLog.SelectionColor = Color.White;
                            }));
                            logger.tfailed.Add(new TestFailed { TestName = "Firmware Check : ", Value = values[4], Result = "FAIL", ErrorCodes = "TT101" });
                            IsFirmwareValid = false;
                        }
                        else
                        {
                            this.Invoke(new MethodInvoker(delegate {
                                testLog.AppendText("Slot#" + activeSlot + " FW version is " + values[4] + "\n");
                            }));
                            IsFirmwareValid = true;
                        }
                    }
                }
                #endregion ConnectToRoot
                stream.WriteLine("ushell");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");



                Thread.Sleep(60000);

                stream.WriteLine("boardAntPowShow");
                Thread.Sleep(1000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("exit");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("exit");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("exit");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("exit");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                timeCheck = DateTime.Now.AddMinutes(2);
                stream.WriteLine("./conn dmp");
                Thread.Sleep(500);
                while (!reader.Contains("sapp@uadpf") && DateTime.Now < timeCheck)
                {
                    Thread.Sleep(10000);
                    reader += stream.Read();
                }
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("nrconfd_cli -u vsmuser");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("set paginate false");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                #region FullPowerCommands
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Sending the full power commands\n");
                }));
                string[] fullPowerCommands = null;
                if (Model == "LOLO" || Model == "FAT LOLO")
                {
                    switch (RUID)
                    {
                        case "700":
                            fullPowerCommands = vZ_LOLO.full700;
                            break;
                        case "701":
                            fullPowerCommands = vZ_LOLO.full701;
                            break;
                        case "702":
                            fullPowerCommands = vZ_LOLO.full702;
                            break;
                        case "800":
                            fullPowerCommands = vZ_LOLO.full800;
                            break;
                        case "801":
                            fullPowerCommands = vZ_LOLO.full801;
                            break;
                            /*case "802":
                                fullPowerCommands = vZ_LOLO.full802;
                                break;*/
                    }
                }
                else if (Model == "PCS")
                {
                    switch (RUID)
                    {
                        case "900":
                            fullPowerCommands = vZ_PCS.full900;
                            break;
                        case "901":
                            fullPowerCommands = vZ_PCS.full901;
                            break;
                        case "902":
                            fullPowerCommands = vZ_PCS.full902;
                            break;
                        case "703":
                            fullPowerCommands = vZ_PCS.full703;
                            break;
                    }
                }
                foreach (string fullPowerCommand in fullPowerCommands)
                {
                    stream.WriteLine(fullPowerCommand);
                    Thread.Sleep(1000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                }

                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Commands sent. Testing Continues at " + DateTime.Now.AddMinutes(3).ToString("hh:mm tt") + "\n");
                }));
                for (int i = 3; i > 0; i--)
                {
                    Thread.Sleep(60000);
                    stream.WriteLine("");
                }
                #endregion FullPowerCommands
                if (!client.IsConnected)
                {
                    client.Connect();
                    stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                    Thread.Sleep(2000);
                    timeCheck = DateTime.Now.AddMinutes(2);
                    stream.WriteLine("./conn dmp");
                    Thread.Sleep(500);
                    while (!reader.Contains("sapp@uadpf") && DateTime.Now < timeCheck)
                    {
                        Thread.Sleep(10000);
                        reader += stream.Read();
                    }
                    Thread.Sleep(1000);
                    stream.WriteLine("nrconfd_cli -u vsmuser");
                    Thread.Sleep(1000);
                    stream.WriteLine("set paginate false");
                    Thread.Sleep(1000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                }
                #region RetCommand
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Looking for RET Signal\n");
                }));
                stream.WriteLine("show table managed-element hardware-management o-ran-radio-unit o-ran-radio-unit-info " + RUID + " antenna-line-device hdlc-state");
                Thread.Sleep(30000);
                stream.WriteLine("");
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                if (reader.Contains("up"))
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("RET Signal is detected\n");
                    }));
                }
                else
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("RET Signal is not detected\n");
                        testLog.SelectionColor = Color.White;
                    }));
                }
                Thread.Sleep(10000);
                #endregion RetCommand

                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Looking for MAC Addresses\n");
                }));
                #region GetMacAddresses
                stream.WriteLine("exit");
                Thread.Sleep(1000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("exit");
                Thread.Sleep(1000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                timeCheck = DateTime.Now.AddMinutes(2);
                stream.WriteLine("./conn rmp");
                Thread.Sleep(500);
                while (!reader.Contains("sapp@uadpf") && DateTime.Now < timeCheck)
                {
                    Thread.Sleep(10000);
                    reader += stream.Read();
                }
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("rm -rf /home/sapp/.ssh/known_hosts");
                Thread.Sleep(1000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("ssh user@" + sshIP);
                Thread.Sleep(5000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                sshCountDown = 3;
                while (!reader.Contains("fingerprint") && sshCountDown > 0)
                {
                    sshCountDown--;
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Failed to SSH Into Unit. Retries left: " + sshCountDown + "\n");
                    }));
                    Thread.Sleep(30000);
                    stream.WriteLine("ssh user@" + sshIP);
                    Thread.Sleep(5000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                }
                if (sshCountDown == 0)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("Failed to SSH Into Unit. Ending test\n");
                        testLog.SelectionColor = Color.White;
                    }));
                    goto EndTest;
                }
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Successful SSH into unit\n");
                }));
                stream.WriteLine("yes");
                Thread.Sleep(1000);
                reader = stream.Read();
                if (!reader.Contains("computer system and network"))
                {
                    client.Disconnect();
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("Infinite y of death. Please wait 3 minutes or until 3 or 4 LEDs are blinking\n");
                        testLog.SelectionColor = Color.White;
                    }));
                    goto EndTest;
                }
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                Thread.Sleep(200);
                stream.WriteLine("REDACTED_PASSWORD");
                Thread.Sleep(12000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                int permissionCountDown = 3;
                while (reader.Contains("Permission denied, please try again.") && permissionCountDown > 0)
                {
                    permissionCountDown--;
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Permission denied. Number of retries left: " + permissionCountDown + "\n");
                    }));
                    stream.WriteLine("");
                    Thread.Sleep(12000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                    stream.WriteLine("");
                    Thread.Sleep(1000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                    stream.WriteLine("ssh user@" + sshIP);
                    Thread.Sleep(5000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                    stream.WriteLine("REDACTED_PASSWORD");
                    Thread.Sleep(12000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                }
                if (permissionCountDown == 0)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("Permission denied. Ending test\n");
                        testLog.SelectionColor = Color.White;
                    }));
                    goto EndTest;
                }

                stream.WriteLine("su -");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("REDACTED_PASSWORD");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("getinv");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                string[] MacLines = reader.Split("\r\n");
                int MacAddress = 0;
                foreach (string line in MacLines)
                {
                    string[] values = line.Split(new String[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                    if (line.Contains("HW SN"))
                    {
                        if (values[3] != scannedSN.Text)
                        {
                            this.Invoke(new MethodInvoker(delegate {
                                testLog.SelectionColor = Color.Red;
                                testLog.AppendText("Scanned serial number does not match internal serial number.\rInternal SN is " + values[3] + "\n");
                                testLog.SelectionColor = Color.White;

                            }));
                        }
                        else
                        {
                            this.Invoke(new MethodInvoker(delegate {
                                testLog.AppendText("Scanned serial number matches internal serial number\n");

                            }));
                        }
                    }
                    if (line.Contains("MAC Address [ 0]") && values[5] != "ff:ff:ff:ff:ff:ff" || line.Contains("MAC Address [ 1]") && values[5] != "ff:ff:ff:ff:ff:ff")
                    {
                        MacAddress++;
                    }
                    else if (line.Contains("MAC Address [ 0]") && values[5] == "ff:ff:ff:ff:ff:ff" || line.Contains("MAC Address [ 1]") && values[5] == "ff:ff:ff:ff:ff:ff")
                    {
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.SelectionColor = Color.Red;
                            testLog.AppendText("Mac Address [ " + MacAddress + "] failed with an address of " + values[5]);
                            testLog.SelectionColor = Color.White;
                        }));
                        MacAddress++;
                    }
                }
                #endregion GetMacAddresses


                stream.WriteLine("ifconfig");
                Thread.Sleep(2000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                string csvIP = GetCSVAddress(reader);
                if (csvIP == "")
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Failed to get address for CSV logging\n");
                    }));
                }
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Checking slots for firmware\n");
                }));
                stream.WriteLine("printenv");
                Thread.Sleep(2000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("gettail 0");
                Thread.Sleep(2000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("disfilever /mnt/storage/slot_1/pkg/bin/*");
                Thread.Sleep(10000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("disfilever /mnt/storage/slot_2/pkg/bin/*");
                Thread.Sleep(10000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                if (!client.IsConnected)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Connection lost. Attempting to log into root\n");

                    }));
                    client.Connect();
                    stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                    Thread.Sleep(2000);
                    timeCheck = DateTime.Now.AddMinutes(2);
                    stream.WriteLine("./conn rmp");
                    Thread.Sleep(500);
                    while (!reader.Contains("sapp@uadpf") && DateTime.Now < timeCheck)
                    {
                        Thread.Sleep(10000);
                        reader += stream.Read();
                    }
                    stream.WriteLine("rm -rf /home/sapp/.ssh/known_hosts");
                    Thread.Sleep(1000);
                    stream.WriteLine("ssh user@" + sshIP);
                    Thread.Sleep(5000);
                    reader = stream.Read();
                    if (reader.Contains("No route to host"))
                    {
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.SelectionColor = Color.Red;
                            testLog.AppendText("Unable to SSH into unit. Failing test\n");
                            testLog.SelectionColor = Color.White;

                        }));
                        goto EndTest;
                    }
                    stream.WriteLine("yes");
                    Thread.Sleep(1000);

                    if (!reader.Contains("computer system and network"))
                    {
                        client.Disconnect();
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.AppendText("Infinite y of death. Please wait 3 minutes or until 3 or 4 LEDs are blinking\n");
                        }));
                        goto EndTest;
                    }

                    stream.WriteLine("REDACTED_PASSWORD");
                    Thread.Sleep(12000);
                    reader = stream.Read();
                    while (reader.Contains("Permission denied, please try again."))
                    {
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.SelectionColor = Color.Red;
                            testLog.AppendText("Permission denied. Ending test\n");
                            testLog.SelectionColor = Color.White;
                        }));
                        goto EndTest;
                    }

                    stream.WriteLine("su -");
                    Thread.Sleep(1000);
                    stream.WriteLine("REDACTED_PASSWORD");
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                }

                stream.WriteLine("ushell");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("boardInvtShow");
                Thread.Sleep(1000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                if (!reader.Contains(scannedSN.Text))
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Serial number not detected in boardInvtshow\n");

                    }));
                }

                stream.WriteLine("almsts"); //Need a method for parsing alarms
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                string alarm = StringCleaner(reader, "almsts", "value = 0 = 0x0");
                (AlarmpnotPresent, vswrDetected) = ParseAlarms(alarm, testLog, logger);

                if (!client.IsConnected)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Connection lost. Attempting to log into ushell\n");

                    }));
                    client.Connect();
                    stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                    Thread.Sleep(2000);
                    timeCheck = DateTime.Now.AddMinutes(2);
                    stream.WriteLine("./conn rmp");
                    Thread.Sleep(500);
                    while (!reader.Contains("sapp@uadpf") && DateTime.Now < timeCheck)
                    {
                        Thread.Sleep(10000);
                        reader += stream.Read();
                    }
                    stream.WriteLine("rm -rf /home/sapp/.ssh/known_hosts");
                    Thread.Sleep(1000);
                    stream.WriteLine("ssh user@" + sshIP);
                    Thread.Sleep(5000);
                    reader = stream.Read();
                    if (reader.Contains("No route to host"))
                    {
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.AppendText("Unable to SSH into unit. Failing test\n");

                        }));
                        IsSshConnected = false;
                        logger.tfailed.Add(new TestFailed { TestName = "Connection: ", Value = "NA", Result = "FAIL", ErrorCodes = "TT023" });
                        goto EndTest;
                    }
                    stream.WriteLine("yes");
                    Thread.Sleep(1000);

                    if (!reader.Contains("computer system and network"))
                    {
                        client.Disconnect();
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.AppendText("Infinite y of death. Please wait 3 minutes or until 3 or 4 LEDs are blinking\n");
                        }));
                        goto EndTest;
                    }

                    stream.WriteLine("REDACTED_PASSWORD");
                    Thread.Sleep(12000);
                    if (reader.Contains("Permission denied, please try again."))
                    {
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.AppendText("Permission denied. Ending test\n");
                        }));
                        goto EndTest;
                    }

                    stream.WriteLine("su -");
                    Thread.Sleep(1000);
                    stream.WriteLine("REDACTED_PASSWORD");
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                    stream.WriteLine("su -");
                    Thread.Sleep(1000);
                    stream.WriteLine("ushell");
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                    Thread.Sleep(1000);
                }

                timer.Start();


                stream.WriteLine("boardAntPowShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                txPowerPassed = ParseAntPowReadings(grid, reader, logger, Model);
                int redoSFP = 1;
            CheckSFP:;
                if (!client.IsConnected)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Connection Lost. Reestablishing connection\n");
                    }));
                    client.Connect();
                    stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                    Thread.Sleep(300);

                    stream.WriteLine("./conn rmp");
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("rm -rf /home/sapp/.ssh/known_hosts");
                    Thread.Sleep(500);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("ssh user@" + sshIP);
                    Thread.Sleep(5000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("yes");
                    Thread.Sleep(5000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("REDACTED_PASSWORD");
                    Thread.Sleep(12000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("su -");
                    Thread.Sleep(500);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("REDACTED_PASSWORD");
                    Thread.Sleep(500);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("ushell");
                    Thread.Sleep(500);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Connection reestablished. Continuing test\n");
                    }));
                }

                stream.WriteLine("boardOptShow"); //Parse out the SFP readings
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                sfpPresent = ParseSFPReadings(reader, testLog, Model, logger);
                if (!sfpPresent && redoSFP > 0)
                {
                    MessageBox.Show("Check port L1 for SFP");
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("\n\n");
                    }));
                    redoSFP--;
                    goto CheckSFP;
                }
                if (!client.IsConnected)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Connection Lost. Reestablishing connection\n");
                    }));
                    client.Connect();
                    stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                    Thread.Sleep(300);

                    stream.WriteLine("./conn rmp");
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("rm -rf /home/sapp/.ssh/known_hosts");
                    Thread.Sleep(500);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("ssh user@" + sshIP);
                    Thread.Sleep(5000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("yes");
                    Thread.Sleep(5000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("REDACTED_PASSWORD");
                    Thread.Sleep(12000);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("su -");
                    Thread.Sleep(500);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("REDACTED_PASSWORD");
                    Thread.Sleep(500);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    stream.WriteLine("ushell");
                    Thread.Sleep(500);
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Connection reestablished. Continuing test\n");
                    }));
                }

                stream.WriteLine("boardInfoShow"); //Parse out the RET Loss readings
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                returnLossPassed = ParseRETLOSS(reader, grid, logger);

                stream.WriteLine("sts");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");


                stream.WriteLine("ipshow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");


                stream.WriteLine("abninfo -a");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("boardEnvShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("boardSourceShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("boardAntPowShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("boardTempShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("boardPllShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("boardOptShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("boardPowVersionShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("boardFAShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("boardFAMapShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("boardInfoShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("boardEmacShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("boardVerShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("boardHwShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("boardSfpShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("Alarm_Print 1");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("IRF_Get_Drain_Bias_Voltage_Level 0 1");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("IRF_Get_Drain_Bias_Voltage_Level 1 1");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("IRF_Get_Drain_Bias_Voltage_Level 2 1");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("IRF_Get_Drain_Bias_Voltage_Level 3 1");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("IRF_Get_Drain_Bias_Voltage_Level 4 1");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("IRF_Get_Drain_Bias_Voltage_Level 5 1");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");


                stream.WriteLine("IRF_Get_Drain_Bias_Voltage_Level 6 1");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("IRF_Get_Drain_Bias_Voltage_Level 7 1");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("dpdsts");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("pacalsts");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("sts 100");
                Thread.Sleep(2000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("almsts");
                Thread.Sleep(2000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                alarm = StringCleaner(reader, "almsts", "value = 0 = 0x0");
                (AlarmpnotPresent, vswrDetected) = ParseAlarms(alarm, testLog, logger);

                if (!client.IsConnected)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Connection lost. Attempting to log into ushell\n");

                    }));
                    client.Connect();
                    stream = client.CreateShellStream("xterm", 80, 50, 1024, 1024, 1024);
                    Thread.Sleep(2000);
                    timeCheck = DateTime.Now.AddMinutes(2);
                    stream.WriteLine("./conn rmp");
                    while (!reader.Contains("sapp@uadpf") && DateTime.Now < timeCheck)
                    {
                        reader += stream.Read();
                    }
                    stream.WriteLine("rm -rf /home/sapp/.ssh/known_hosts");
                    Thread.Sleep(1000);
                    stream.WriteLine("ssh user@" + sshIP);
                    Thread.Sleep(5000);
                    reader = stream.Read();
                    if (reader.Contains("No route to host"))
                    {
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.AppendText("Unable to SSH into unit. Failing test\n");

                        }));
                        IsSshConnected = false;
                        logger.tfailed.Add(new TestFailed { TestName = "Connection: ", Value = "NA", Result = "FAIL", ErrorCodes = "TT023" });
                        goto EndTest;
                    }
                    stream.WriteLine("yes");
                    Thread.Sleep(1000);

                    if (!reader.Contains("computer system and network"))
                    {
                        client.Disconnect();
                        this.Invoke(new MethodInvoker(delegate {
                            testLog.AppendText("Infinite y of death. Please wait 3 minutes or until 3 or 4 LEDs are blinking\n");
                        }));
                        goto EndTest;
                    }

                    stream.WriteLine("REDACTED_PASSWORD");
                    Thread.Sleep(12000);
                    reader = stream.Read();
                    if (reader.Contains("Permission denied, please try again."))
                    {

                        this.Invoke(new MethodInvoker(delegate {
                            testLog.AppendText("Permission denied. Ending test\n");
                        }));
                        goto EndTest;
                    }

                    stream.WriteLine("su -");
                    Thread.Sleep(1000);
                    stream.WriteLine("REDACTED_PASSWORD");
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                    stream.WriteLine("su -");
                    Thread.Sleep(1000);
                    stream.WriteLine("ushell");
                    reader = stream.Read();
                    File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                    Thread.Sleep(1000);
                }

                timer.Start();

                stream.WriteLine("boardInvtShow");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("console sts");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                stream.WriteLine("exit");
                Thread.Sleep(500);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");

                reader = string.Empty;
                bool secureIsTrue = false;
                stream.WriteLine("cat /proc/sys/platform/etc/secureboot_status");
                Thread.Sleep(1000);
                reader = stream.Read();
                File.AppendAllText(logfile, "\nElapsed Time: " + timeStamp.Text + "\n" + reader + "\n");
                foreach (string line in reader.Split("\r\n"))
                {
                    if (line.Length == 1 && line.Contains("1"))
                    {
                        secureIsTrue = true;
                    }
                }
                if (secureIsTrue)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText("Secure boot is enabled\n");
                    }));
                }
                else
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("Secure boot is not enabled\n");
                        testLog.SelectionColor = Color.White;
                    }));
                }
                client.Disconnect();
                if (!RecentCsvExists(scannedSN.Text, Model))
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText(
                        "\nNo CSV log found within the last 14 days. Retrieving CSV log...\n");
                    }));

                    GetCSVLog(scannedSN.Text, DU_IP, sshIP, csvIP, logfile, Model, testLog, 3, vswrDetected);
                }
                else
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.AppendText(
                        "\nRecent CSV log found. Skipping CSV retrieval.\n");
                    }));

                }

            EndTest:;
                #region OLP Data Entry 

                if (txPowerPassed == true && returnLossPassed == true && AlarmpnotPresent == true && IsFirmwareValid == true && IsSshConnected == true && sfpPresent == true)
                {
                    logger.tfailed.Add(new TestFailed { TestName = "Firmware Check: ", Value = Firmwareactive, Result = "PASS", ErrorCodes = "'NA" });
                    logger.tfailed.Add(new TestFailed { TestName = "Connection: ", Value = "NA", Result = "PASS", ErrorCodes = "NA" });
                    logger.tfailed.Add(new TestFailed { TestName = "Tx Power 1 TO 8 : ", Result = "PASS", ErrorCodes = "NA" });
                    logger.tfailed.Add(new TestFailed { TestName = "RSSI 1 TO 8 : ", Result = "PASS", ErrorCodes = "NA" });
                    logger.tfailed.Add(new TestFailed { TestName = "Ret Loss  1 TO 8 : ", Result = "PASS", ErrorCodes = "NA" });
                    logger.tfailed.Add(new TestFailed { TestName = "SFP Test :", Result = "PASS", ErrorCodes = "NA" });
                    logger.tfailed.Add(new TestFailed { TestName = "Alarm : ", Result = "PASS", ErrorCodes = "NA" });

                    logger.tlog.OverallResult = "PASS";
                }
                else
                {
                    logger.tlog.OverallResult = "FAIL";
                }

                if (radioInitial_LOLO.Checked && (Model == "LOLO" || Model == "FAT LOLO"))
                {
                    logger.tlog.WorkStation = "ORAN Initial";
                }
                else if (radioSystem_LOLO.Checked && (Model == "LOLO" || Model == "FAT LOLO"))
                {
                    logger.tlog.WorkStation = "ORAN System";
                }
                if (radioInitial_PCS.Checked && Model == "PCS")
                {
                    logger.tlog.WorkStation = "ORAN Initial";
                }
                else if (radioSystem_PCS.Checked && Model == "PCS")
                {
                    logger.tlog.WorkStation = "ORAN System";
                }

                logger.tlog.SerialNumber = serialNumber;
                logger.tlog.DateTime = DateTime.Now.ToString();
                //logger.tlog.SlotID = slot.ToString();
                logger.tlog.Firmware = Firmwareactive.ToString();
                logger.tlog.Model = Model;
                logger.tlog.Locations = "Facility 1";

                bool Ftp_FileisCopied = logger.WriteToLog(serialNumber);
                logger.tfailed.Clear();

                if (Ftp_FileisCopied == true)
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.LightBlue;
                        testLog.AppendText("Json Copied to the Server...!\n");
                        testLog.SelectionColor = Color.White;
                    }));
                }
                else
                {
                    this.Invoke(new MethodInvoker(delegate {
                        testLog.SelectionColor = Color.Red;
                        testLog.AppendText("Unable to copy the file to the Server...!\n");
                        testLog.SelectionColor = Color.White;
                    }));

                }


                #endregion

                timer.Stop();
                if (client.IsConnected)
                {
                    client.Disconnect();
                }
                this.Invoke(new MethodInvoker(delegate {
                    testLog.AppendText("Test has ended at " + DateTime.Now.ToString("hh:mm tt"));
                }));
                File.AppendAllText(logfile, "\nLog Ended\r**********************************************************");
                try
                {
                    if (!string.IsNullOrEmpty(LogFileList[1]))
                    {
                        string destDir = IOPath.GetDirectoryName(LogFileList[1]);
                        Directory.CreateDirectory(destDir);
                        File.Copy(logfile, LogFileList[1], true);
                    }
                }
                catch (Exception ex)
                {
                    Task.Run(() => { MessageBox.Show("Failed to copy log to T: drive\n" + ex.ToString()); });
                }
                /* try {
                     using (SftpClient sftpClient = new SftpClient("sftp.example.com", 22, "REDACTED_USER", "REDACTED_PASSWORD")) {
                         sftpClient.Connect();
                         using (var fileStream = new FileStream(logfile, FileMode.Open)) {
                             sftpClient.UploadFile(fileStream, Listlogfile[1], true);
                         }
                         sftpClient.Disconnect();
                     }

                 }
                 catch (Exception ex) {
                     Task.Run(() => { MessageBox.Show("Failed to upload to remote server\n" + ex.ToString()); });
                 }
                 // Run chmod via SSH exec (captures exit status and stderr/stdout)
                 try {
                     SetRemotePermissions(Listlogfile[1]);
                 }
                 catch (Exception ex) {
                     Task.Run(() => MessageBox.Show("Upload / chmod test failed:\n" + ex.ToString()));
                 }*/
            }
        }
        #endregion MainTest

        #region Menu and Toolbar Events

        private void vDUSettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            VDU_AssignmentForm vDU_AssignmentForm = new VDU_AssignmentForm();
            vDU_AssignmentForm.Show();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About about = new About();
            about.Show();
        }

        private void gUISettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Setting setting = new Setting();
            setting.Show();
        }

        #endregion Menu and Toolbar Events

        #region Debug and Manual Test Buttons

        private void button1_Click(object sender, EventArgs e)
        {
            //bool exists =
            //        RecentCsvExists(
            //            "SAMPLE0002",
            //            "PCS");

            //MessageBox.Show(
            //    exists
            //    ? "Recent CSV Found"
            //    : "No Recent CSV Found");

            string output = SSH_IP("mplane-info mplane-interfaces mplane-ipv6 REDACTED_IPV6", "703");

            MessageBox.Show(output);
        }


        private void button2_Click_1(object sender, EventArgs e)
        {
            // Hardcoded sample serial number for manually exercising LogHandler with known-good
            // test data, without needing a live unit connected. (Already inside the Debug and
            // Manual Test Buttons region, noted here too for clarity.)
            string serialNumber = "SAMPLE0003";
            string input = "almsts\r\n------------------------------------------------------\r\n\r\n[No] AlmName          : AntId(PaId)| I_FA| E_FA| CurSts| SubCodeType \r\n\r\n[ 1] \u001b[31mShutDown         \u001b[0m:    7 (   6)|    0|  N/A|  Occur|     PATH   \r\n\r\n[ 1] \u001b[31mShutDown         \u001b[0m:    8 (   7)|    0|  N/A|  Occur|     PATH   \r\n\r\n[ 9] \u001b[31mLowGain          \u001b[0m:    7 (   6)|    0|  N/A|  Occur|     PATH   \r\n\r\n[ 9] \u001b[31mLowGain          \u001b[0m:    8 (   7)|    0|  N/A|  Occur|     PATH   \r\n\r\n[No] AlmName          :    DeviceId| I_FA| E_FA| CurSts| SubCodeType \r\n\r\n[19] \u001b[33mOptRxLOS         \u001b[0m:           1|    0|  N/A|  Occur|     CPRI   \r\n\r\n[28] \u001b[33mUDA              \u001b[0m:           0|    0|  N/A|  Occur|      UDA   \r\n\r\n[28] \u001b[33mUDA              \u001b[0m:           1|    0|  N/A|  Occur|      UDA   \r\n\r\n[28] \u001b[33mUDA              \u001b[0m:           2|    0|  N/A|  Occur|      UDA   \r\n\r\n[28] \u001b[33mUDA              \u001b[0m:           3|    0|  N/A|  Occur|      UDA   \r\n\r\n[53] \u001b[33mGroupTxShutdown  \u001b[0m:           0|    0|  N/A|  Occur|    SYSTEM  \r\n\r\n------------------------------------------------------\r\n\r\nvalue = 0 = 0x0\r\n";

            LogHandler logger = new LogHandler();
            ParseAlarms(input, testLog_PCS, logger);

            input = "boardAntPowShow\r\n**********************Path Power Domain***************************\r\n      Item| Path 0| Path 1| Path 2| Path 3| Path 4| Path 5| Path 6| Path 7|\r\n  TxAntSum|  45.67|  45.66|  45.66|  45.64|  45.68|  45.68|  -8.81|  -9.50|\r\n RxAntFa00| -98.26| -98.43| -97.96| -98.48|-101.40|-101.72|-100.97|-101.27|\r\n RxAntFa01| -99.16| -98.90| -98.95| -99.16|-101.56|-101.98|-101.30|-101.56|\r\n RxAntFa02|-110.79|-110.79|-110.79|-110.79|-110.79|-110.79|-110.79|-110.79|\r\n RxAntFa03|-110.79|-110.79|-110.79|-110.79|-110.79|-110.79|-110.79|-110.79|\r\n RxAntFa04|-110.79|-110.79|-110.79|-110.79|-110.79|-110.79|-110.79|-110.79|\r\n RxAntFa05|-110.79|-110.79|-110.79|-110.79|-110.79|-110.79|-110.79|-110.79|\r\n RxAntFa06|-110.79|-110.79|-110.79|-110.79|   0.00|   0.00|   0.00|   0.00|\r\n RxAntFa07|-110.79|-110.79|-110.79|-110.79|   0.00|   0.00|   0.00|   0.00|\r\n RxAntFa08|-110.79|-110.79|-110.79|-110.79|   0.00|   0.00|   0.00|   0.00|\r\n RxAntFa09|-110.79|-110.79|-110.79|-110.79|   0.00|   0.00|   0.00|   0.00|\r\nvalue = 0 = 0x0\r\n\r\nUShell >\r\n";

            //ParseAntPowReadings(dataGridView1, input, logger, Model);

            ParseRETLOSS("Ret Loss|  25.00 dB|  25.00 dB|  25.00 dB|  25.00 dB|  25.00 dB|  19.09 dB|        NA|        NA|", dataGridView1, logger);

            logger.tlog.OverallResult = "FAIL";

            logger.tlog.WorkStation = "ORAN System/Initial";
            logger.tlog.SerialNumber = serialNumber;
            logger.tlog.DateTime = DateTime.Now.ToString();
            //logger.tlog.SlotID = slot.ToString();
            logger.tlog.Firmware = "24.A.41";
            logger.tlog.Model = "PCS";
            logger.tlog.Locations = "Facility 1";

            logger.WriteToLog(serialNumber);
        }

        #endregion Debug and Manual Test Buttons

        #region RUID Label Click Handlers


        private void PCSRUIDLabel_Click(object sender, EventArgs e)
        {
            if (scannedSN_PCS.Enabled == true)
            {
                switch (PCSRUIDLabel.Text)
                {
                    case "703":
                        PCSIPLabel.Text = "REDACTED_IP";
                        PCSRUIDLabel.Text = "900";
                        break;
                    case "900":
                        PCSIPLabel.Text = "REDACTED_IP";
                        PCSRUIDLabel.Text = "901";
                        break;
                    case "901":
                        PCSIPLabel.Text = "REDACTED_IP";
                        PCSRUIDLabel.Text = "902";
                        break;
                    case "902":
                        PCSIPLabel.Text = "REDACTED_IP";
                        PCSRUIDLabel.Text = "703";
                        break;
                }
            }
        }

        private void LOLORUIDLabel_Click(object sender, EventArgs e)
        {
            if (scannedSN_LOLO.Enabled == true)
            {
                switch (LOLORUIDLabel.Text)
                {
                    case "700":
                        LOLORUIDLabel.Text = "701";
                        break;
                    case "701":
                        LOLORUIDLabel.Text = "702";
                        break;
                    case "702":
                        LOLORUIDLabel.Text = "700";
                        break;
                    case "800":
                        LOLORUIDLabel.Text = "801";
                        break;
                    case "801":
                        LOLORUIDLabel.Text = "800";
                        break;
                }
            }
        }

        #endregion RUID Label Click Handlers

    }
}