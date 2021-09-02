using System.Web;
using System.Data;
//using System.Web.Security;
//using System.Web.UI;
//using System.Web.UI.WebControls;
//using System.Web.UI.HtmlControls;
//using Microsoft.SharePoint;
using System.Text;
using System.Collections.Generic;
using System;
using System.Globalization;
using System.ComponentModel;
using System.IO;
using System.Xml;
using System.Collections;
using System.Configuration;
using System.Diagnostics;
using System.Security;
using System.Security.Policy;
using System.Security.Principal;
using System.Security.Permissions;
using System.Runtime.InteropServices;

namespace Common
{
    public class Logging
    {

        #region class properties
        public static string LOG_DIRECTORY { get; set; }
        //public string LOG_DIRECTORY { get; set; }
        public const string Info = "INFO";
        public const string Error = "ERROR";
        private string logFolderPath = LOG_DIRECTORY;

        private string LogFolderPath
        {
            get
            {
                return logFolderPath;
            }
            set
            {
                if (Directory.Exists(value))
                {
                    logFolderPath = value;
                    if (logFolderPath.EndsWith("\\"))
                    {
                        logFolderPath = logFolderPath.Remove(logFolderPath.Length);
                    }
                }
                else
                {
                    throw new DirectoryNotFoundException();
                }
            }
        }
        private string logFileName = "";
        private string LogFileName
        {
            get
            {
                return logFileName;
            }
            set
            {
                logFileName = value;
            }
        }
        private string LogFilePath
        {
            get
            {
                return this.LogFolderPath + "\\" + this.LogFileName;
            }
        }
        private static Logging _log;
        #endregion

        public static Logging GetInstance()
        {
            if (_log == null)
                _log = new Logging(LOG_DIRECTORY, "Log_", true, false);
            if(!_log.logFolderPath.Equals(LOG_DIRECTORY))
                _log = new Logging(LOG_DIRECTORY, "Log_", true, false);
            return _log;
        }

        #region Constructor
        /// <summary>
        /// Create a HTMLFileLogging object
        /// </summary>
        /// <param name="folderPath">The path of the folder that will hold the log file (no file name)</param>
        /// <param name="fileName">The name of the file to create</param>
        /// <param name="createPath">When this is set to true and the folder does not exist, the code will create the folder.</param>
        /// <param name="deleteFile">When this is set to true and the file exists, the code will delete the file and create a new one</param>
        private Logging(string folderPath, string fileName, bool createPath, bool deleteFile)
        {
            this.LogFileName = fileName + DateString() + ".log";

            if (createPath && !Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            this.LogFolderPath = folderPath;
            if (File.Exists(this.LogFilePath) && deleteFile)
            {
                File.Delete(this.LogFilePath);
            }
        }

        #endregion

        #region custom code
        /// <summary>
        /// Writes a string to the log file. 
        /// </summary>
        /// <param name="message">a string to write. supports html tags.</param>
        public void WriteToLogFile(string level, string message)
        {
            try
            {
                //string rowFormat = "<table border=0><tr><td nowrap style=\"font-size:x-small;width:200px\" valign='top'><date>{0}</date> <time>{1}</time></td><td  style=\"font-size:x-small;width:450px\"> <message>{2}</message></td></tr></table>";
                StreamWriter sw = new StreamWriter(this.LogFilePath, true);
                string logMesg = String.Format(DateTime.Now.ToString(), "\t", level, "\t", message, "\n");
                logMesg = string.Concat(DateTime.Now.ToString(), "\t", level, "\t", message, "\n");
                sw.WriteLine(logMesg);
                sw.Flush();
                sw.Close();
            }
            catch
            {
            }
        }

        public string DateString()
        {
            string dateString = string.Empty;

            if (DateTime.Now.Month < 10)
                dateString += "0";
            dateString += DateTime.Now.Month;

            if (DateTime.Now.Day < 10)
                dateString += "0";
            dateString += DateTime.Now.Day;

            dateString += DateTime.Now.Year.ToString();

            dateString += "_";

            if (DateTime.Now.Hour < 10)
                dateString += "0";
            dateString += DateTime.Now.Hour;

            dateString += "_";

            if (DateTime.Now.Minute < 10)
                dateString += "0";
            dateString += DateTime.Now.Minute;

            dateString += "_";

            if (DateTime.Now.Second < 10)
                dateString += "0";
            dateString += DateTime.Now.Second;


            return dateString;
        }

        #endregion

    }
}