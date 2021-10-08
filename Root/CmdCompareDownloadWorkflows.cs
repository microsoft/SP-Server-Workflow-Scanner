using System;
using System.IO;
using System.Management.Automation;
using Common;
namespace Root
{
    public class CmdCompareDownloadWorkflows : PSCmdlet
    {
        private string logFolderPath;
        public DirectoryInfo logFolder;

        [Parameter(Mandatory = true, HelpMessage = @"The path where workflows are downloaded to for analyzing (e.g. c:\temp\workflows")]
        public string OutputDirectory { get; set; }

        protected override void BeginProcessing()
        {
            if (!Directory.Exists(string.Concat(OutputDirectory, @"\Logs")))
            {
                logFolder = System.IO.Directory.CreateDirectory(string.Concat(OutputDirectory, @"\Logs"));
                logFolderPath = logFolder.FullName;
                Logging.LOG_DIRECTORY = logFolderPath;
                Logging.GetInstance().WriteToLogFile(Logging.Info, "Log folder created");

            }
            else
            {
                logFolderPath = string.Concat(OutputDirectory, @"\Logs");
                Logging.LOG_DIRECTORY = logFolderPath;
            }

            base.BeginProcessing();
        }

        protected override void EndProcessing()
        {
            base.EndProcessing();
        }

        protected override void ProcessRecord()
        {
            Console.WriteLine(System.Environment.NewLine);
            Host.UI.WriteLine(ConsoleColor.Cyan, Host.UI.RawUI.BackgroundColor, "Beginning to analyze and compare SharePoint Workflows to Power Automate features.. ");
            Console.WriteLine(System.Environment.NewLine);
            Host.UI.WriteLine(ConsoleColor.Cyan, Host.UI.RawUI.BackgroundColor, "Please standby. This may take some time.. ");
            Console.WriteLine(System.Environment.NewLine);
            Logging.GetInstance().WriteToLogFile(Logging.Info, "Beginning to analyze and compare SharePoint Workflows to Power Automate features.. ");
            CompareWorkflows cwf = new CompareWorkflows();
            cwf.CompareWorkflowsToPowerAutomate(OutputDirectory);
        }
    }
}
