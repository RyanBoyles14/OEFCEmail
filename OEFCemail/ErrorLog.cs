using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;

namespace OEFCemail
{

    //https://cybarlab.com/save-error-log-in-text-file-in-c-sharp
    class ErrorLog
    {
        public bool WriteErrorLog(string LogMessage)
        {
            bool Status = false;
            string LogDirectory = ConfigurationManager.AppSettings["LogDirectory"].ToString();

            DateTime CurrentDateTime = DateTime.Now;
            CheckCreateLogDirectory(LogDirectory);
            string logLine = BuildLogLine(CurrentDateTime, LogMessage);
            LogDirectory = (LogDirectory + "Log_" + LogTime(DateTime.Now) + ".txt");

            lock (typeof(ErrorLog))
            {
                StreamWriter oStreamWriter = null;
                try
                {
                    oStreamWriter = new StreamWriter(LogDirectory, true);
                    oStreamWriter.WriteLine(logLine);
                    Status = true;
                }
                catch
                {

                }
                finally
                {
                    if (oStreamWriter != null)
                    {
                        oStreamWriter.Close();
                    }
                }
            }
            return Status;
        }


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


        private string BuildLogLine(DateTime CurrentDateTime, string LogMessage)
        {
            StringBuilder loglineStringBuilder = new StringBuilder();
            loglineStringBuilder.Append(LogTime(CurrentDateTime));
            loglineStringBuilder.Append(" \t");
            loglineStringBuilder.Append(LogMessage);
            return loglineStringBuilder.ToString();
        }

        private string LogTime(DateTime CurrentDateTime)
        {
            return CurrentDateTime.ToString("dd-MM-yyyy HH:mm:ss");
        }
    }
}
