using System;
using System.Configuration;
using System.IO;
using System.Text;
using System.Windows.Forms;
using JR.Utils.GUI.Forms;

namespace OEFCemail
{

    //https://cybarlab.com/save-error-log-in-text-file-in-c-sharp
    public class ErrorLog
    {
        // get preferred directory to save logs in from AppSettings
        private readonly string LogDirectory = ConfigurationManager.AppSettings["LogDirectory"].ToString();
        private static string file;
        private static string time;
        private bool isException = false;

        public ErrorLog()
        {
            CheckCreateLogDirectory(LogDirectory);
            time = LogTime(DateTime.Now);
            file = Path.Combine(LogDirectory, "Log_" + time + ".txt");
        }

        public bool IsException
        {
            get { return isException; }
        }

        public event EventHandler<SendErrorReportEventArgs> SendErrorReport;

        public class SendErrorReportEventArgs : EventArgs
        {
            public string File { get; set; }
            public string Time { get; set; }
        }

        protected virtual void OnSendErrorReport(SendErrorReportEventArgs e)
        {
            EventHandler<SendErrorReportEventArgs> handler = SendErrorReport;
            handler?.Invoke(this, e);
        }

        public bool WriteErrorLog(string logMessage)
        {
            isException = true;

            bool status = false;
            string logLine = BuildLogLine(DateTime.Now, logMessage);
            
            // write to log file
            lock (typeof(ErrorLog))
            {
                StreamWriter oStreamWriter = null;
                try
                {
                    oStreamWriter = new StreamWriter(file, true);
                    oStreamWriter.WriteLine(logLine);
                    status = true;
                }
                finally
                {
                    if (oStreamWriter != null)
                    {
                        oStreamWriter.Close();
                    }
                }
            }

            FlexibleMessageBox.Show(
                $"Exception thrown with message: {logMessage}\r" +
                "The process will be terminated.\r"
               , "Exception Thrown"
            );

            return status;
        }

        public void ErrorReport()
        {
            DialogResult result = FlexibleMessageBox.Show(
                "Exception(s) caught: Would you like to automatically send an error report?\r" +
                "(Note: this will not forward the email you wanted to intake, in case of confidential information.)"
               , "Exception Thrown", MessageBoxButtons.YesNoCancel
            );

            bool sendReport = result == DialogResult.Yes;

            // After writing the log, invoke the error handler in ThisAddIn if the user wants to send an error report
            // Only ThisAddIn can send emails via Outlook
            if (sendReport && SendErrorReport != null)
            {
                SendErrorReportEventArgs args = new SendErrorReportEventArgs
                {
                    File = file,
                    Time = time
                };
                SendErrorReport?.Invoke(this, args);
            }
        }

        // Check for directory to save log files
        // if none, create it
        private bool CheckCreateLogDirectory(string LogPath)
        {
            bool loggingDirectoryExists = false;
            DirectoryInfo oDirectoryInfo = new DirectoryInfo(LogPath);
            if (oDirectoryInfo.Exists)
            {
                loggingDirectoryExists = true;
            }
            else
            {
                try
                {
                    Directory.CreateDirectory(LogPath);
                    loggingDirectoryExists = true;
                }
                catch
                {
                    // Logging failure
                }
            }
            return loggingDirectoryExists;
        }

        // return string of error log to save in file
        private string BuildLogLine(DateTime CurrentDateTime, string LogMessage)
        {
            StringBuilder loglineStringBuilder = new StringBuilder();
            loglineStringBuilder.Append(LogTime(CurrentDateTime));
            loglineStringBuilder.Append(" \t");
            loglineStringBuilder.Append(LogMessage);
            return loglineStringBuilder.ToString();
        }

        // return formatted timestamp for file naming and timestamping error logs
        private string LogTime(DateTime CurrentDateTime)
        {
            return CurrentDateTime.ToString("yyyy-MM-dd_HH.mm.ss");
        }
    }
}
