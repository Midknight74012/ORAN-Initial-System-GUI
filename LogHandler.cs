using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Newtonsoft.Json;




namespace ORAN_Initial_System_GUI
{
    public class TestLog
    {
        public string WorkStation;
        public string SerialNumber;
        public string DateTime;
        public string Model;
        public string Locations;
        public string Firmware;
        public string OverallResult;
        public List<TestFailed> TestDetail;
    }

    public class TestFailed
    {
        public string TestName;
        public string Result;
        public dynamic Value;
        public string ErrorCodes;
    }

    public class LogHandler
    {
        public TestLog tlog = new TestLog();
        public List<TestFailed> tfailed = new List<TestFailed>();

        public bool WriteToLog(string sn)
        {
            bool fileCopied = false;

            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string localJsonPath = Path.Combine(@"C:\JsonLog", sn + "_InitialSystem_" + timestamp + ".json");
            string localTxtPath = Path.Combine(@"C:\Logs", sn + "_InitialSystem_" + timestamp + ".txt");

            tlog.TestDetail = tfailed;
            tlog.DateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            tlog.SerialNumber = sn;

            // Serialize to JSON
            //string JSONResult = JsonConvert.SerializeObject(tlog, Formatting.Indented);
            string JSONResult = JsonConvert.SerializeObject(tlog, Newtonsoft.Json.Formatting.Indented);


            // Ensure directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(localJsonPath));
            Directory.CreateDirectory(Path.GetDirectoryName(localTxtPath));

            // Write JSON log
            File.WriteAllText(localJsonPath, JSONResult);

            // ======== FTP Upload ========
            try
            {
                string ftpServer = "ftp://sftp.example.com";
                string remoteDir = "/production_json";
                string ftpUser = "REDACTED_USER";
                string ftpPass = "REDACTED_PASSWORD";

                if (!ftpServer.EndsWith("/")) ftpServer += "/";
                if (!remoteDir.EndsWith("/")) remoteDir += "/";

                string ftpUrl = ftpServer + remoteDir + Path.GetFileName(localJsonPath);

                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(ftpUrl);
                request.Method = WebRequestMethods.Ftp.UploadFile;
                request.Credentials = new NetworkCredential(ftpUser, ftpPass);

                byte[] fileContents = File.ReadAllBytes(localJsonPath);
                request.ContentLength = fileContents.Length;

                using (Stream requestStream = request.GetRequestStream())
                {
                    requestStream.Write(fileContents, 0, fileContents.Length);
                }

                using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
                {
                    fileCopied = true;
                    Console.WriteLine($"FTP Upload complete, status: {response.StatusDescription}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("FTP upload failed: " + ex.Message);
            }
            // ======== End FTP Upload ========

            TestLog deserialized = JsonConvert.DeserializeObject<TestLog>(JSONResult);
            // Write Text Log
            try
            {
                using (StreamWriter r = File.Exists(localTxtPath) ? File.AppendText(localTxtPath) : File.CreateText(localTxtPath))
                {
                    r.WriteLine("\n\nTELCOM INC.");
                    r.WriteLine("SERIAL NUMBER: " + deserialized.SerialNumber);
                    r.WriteLine("DATE/TIME: " + deserialized.DateTime);
                    r.WriteLine("MODEL: " + deserialized.Model);
                    r.WriteLine("LOCATION: " + deserialized.Locations);
                    r.WriteLine("");

                    foreach (var item in deserialized.TestDetail)
                    {
                        r.WriteLine(item.TestName.ToString() + ":" + item.Result.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Text log writing failed: " + ex.Message);
            }

            // Clear failed tests for next log
            //tfailed.Clear();

            return fileCopied;
        }
    }
}
