using System;
using System.Configuration;
using System.IO;
using System.Text;

namespace OEFCemail
{

    //https://cybarlab.com/save-error-log-in-text-file-in-c-sharp
    class ErrorLog
    {
        //TODO: revise to only have one log file and one instance of ErrorLog, in case of multiple exceptions in one run of EmailSaver

        public bool WriteErrorLog(string LogMessage)
        {
            bool status = false;

            // get preferred directory to save logs in from AppSettings
            string LogDirectory = ConfigurationManager.AppSettings["LogDirectory"].ToString();

            DateTime CurrentDateTime = DateTime.Now;
            CheckCreateLogDirectory(LogDirectory);
            string logLine = BuildLogLine(CurrentDateTime, LogMessage);
            string file = Path.Combine(LogDirectory, "Log_" + LogTime(DateTime.Now) + ".txt");

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
            return status;
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
