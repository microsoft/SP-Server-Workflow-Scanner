using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Management.Automation;
using Common;


namespace Root
{
    /// <summary>
    /// Commandlet to discover the InfoPath forms in 
    /// Onpremise environments that are published to
    /// list and document libraries
    /// </summary>
    /// <returns></returns>   /// 
    ///

    [Cmdlet(VerbsCommon.Get, "WorkflowAssociationsForOnprem")]
    public class CmdGetWorkflowAssociationsForOnprem : PSCmdlet
    {
        private string assessmentScope;
        private static List<string> sitecollectionUrls = new List<string>();

        [Parameter(Mandatory = true, HelpMessage = "Specify the Domain name of the user account")]
        public string DomainName;

        [Parameter(Mandatory = false, ParameterSetName = "Credential", HelpMessage = "Specify the user name and password ")]
        public PSCredential Credential;

        [Parameter(Mandatory = false, HelpMessage = "Specify the file path of a text file containing target site collection URLs")]
        public string SiteCollectionURLFilePath;

       [Parameter(Mandatory = false, HelpMessage = "Specify the URL of the Web Application")]
        public string WebApplicationUrl;

        [Parameter(Mandatory = false, HelpMessage = "Specify the URL of the Site Collection")]
        public string SiteCollectionUrl;
        public bool DownloadWorkflowDefinitions
        {
            get { return downloadWFDefinitions; }
            set { downloadWFDefinitions = value; }
        }
        private bool downloadWFDefinitions;

        [Parameter(Mandatory = true, HelpMessage = @"The path where the Assessment Summary, logs, Workflow definitions are downloaded (if DownloadForms parameter is set to true) for analyzing (e.g. F:\temp\WorkflowDefinitions")]
        public string AssessmentOutputFolder;
        private string logFolderPath;
        public DirectoryInfo logFolder;

        public DataTable dtWorkflowLocations = new DataTable();
        protected override void BeginProcessing()
        {

            if (!Directory.Exists(string.Concat(AssessmentOutputFolder, @"\Logs")))
            {
                logFolder = System.IO.Directory.CreateDirectory(string.Concat(AssessmentOutputFolder, @"\Logs"));
                logFolderPath = logFolder.FullName;
                Logging.LOG_DIRECTORY = logFolderPath;
                Logging.GetInstance().WriteToLogFile(Logging.Info, "Log folder created");

            }
            else
            {
                logFolderPath = string.Concat(AssessmentOutputFolder, @"\Logs");
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
            try
            {

                if (String.IsNullOrEmpty(WebApplicationUrl))
                {
                     if (!String.IsNullOrEmpty(SiteCollectionURLFilePath))
                    {
                        assessmentScope = "SiteCollectionsUrls";
                        BeginToAssess();
                    }
                    else if (String.IsNullOrEmpty(SiteCollectionUrl))
                    {
                        assessmentScope = "Farm";
                        BeginToAssess();
                    }
                    else if(!String.IsNullOrEmpty(SiteCollectionURLFilePath))
                    {
                        assessmentScope = "SiteCollectionsUrls";
                        BeginToAssess();
                    }
                    else
                    {
                        assessmentScope = "SiteCollection";
                        BeginToAssess();
                    }
                }
                else
                {
                    if (!String.IsNullOrEmpty(SiteCollectionUrl))
                    {
                        WriteWarning("Provide either the Web App URL or the Site Collection URL, but not both !");

                    }

                    else
                    {
                        assessmentScope = "WebApplication";
                        BeginToAssess();
                    }
                }

            }
            catch (Exception ex)
            {
                Host.UI.WriteLine(ConsoleColor.DarkRed, Host.UI.RawUI.BackgroundColor, ex.Message);
            }
        }


        protected void BeginToAssess()
        {
            Operations ops = new Operations();
            try
            {
                //New Code Starts
                string userInput = string.Empty;
                Console.WriteLine(System.Environment.NewLine);
                Host.UI.WriteLine(ConsoleColor.Cyan, Host.UI.RawUI.BackgroundColor, "The assessment is scoped to run at " + assessmentScope +
                       " level. Would you like to proceed? [Y] to continue, [N] to abort.");
                var op = this.InvokeCommand.InvokeScript("Read-Host");
                userInput = op[0].ToString().ToLower();

                while (!userInput.Equals("y") && !userInput.Equals("n"))
                {
                    Host.UI.WriteLine(ConsoleColor.Cyan, Host.UI.RawUI.BackgroundColor, "Invalid input. Press [Y] to continue, [N] to abort.");
                    op = this.InvokeCommand.InvokeScript("Read-Host");
                    userInput = op[0].ToString().ToLower();

                }
                if (userInput.Equals("y"))
                {
                    GetWorkflowsforOnPrem objonPrem = new GetWorkflowsforOnPrem();
                    ops.CreateDirectoryStructure(AssessmentOutputFolder);
                    Console.WriteLine(System.Environment.NewLine);
                    Host.UI.WriteLine(ConsoleColor.Yellow, Host.UI.RawUI.BackgroundColor, "Beginning assessment..");

                    if (assessmentScope.Equals("Farm"))
                    {
                        objonPrem.Scope = "Farm";
                        objonPrem.Url = null;
                    }
                    else if (assessmentScope.Equals("WebApplication"))
                    {
                        objonPrem.Scope = "WebApplication";
                        objonPrem.Url = WebApplicationUrl;
                    }
                    else if (assessmentScope.Equals("SiteCollection"))
                    {
                        objonPrem.Scope = "SiteCollection";
                        objonPrem.Url = SiteCollectionUrl;
                    }
                    else if (assessmentScope.Equals("SiteCollectionsUrls"))
                    {
                        objonPrem.Scope = "SiteCollectionsUrls";
                        objonPrem.Url = SiteCollectionURLFilePath;

                    }
                    objonPrem.DownloadPath = AssessmentOutputFolder;
                    //Set Credentials from user entry 
                    if (Credential != null)
                    {
                        objonPrem.Credential = Credential;
                    } 
                    else
                    { 
                    }
                    
                    // run the workflow scan
                    dtWorkflowLocations = objonPrem.Execute();
                    //Save the CSV file
                    string csvFilePath = string.Concat(AssessmentOutputFolder, ops.summaryFolder, ops.summaryFile);
                    ops.WriteToCsvFile(dtWorkflowLocations, csvFilePath);
                }
                else if (userInput.Equals("n"))
                {
                    Host.UI.WriteLine(ConsoleColor.Cyan, Host.UI.RawUI.BackgroundColor, "Operation aborted as per your input !");
                }              
            }
            catch (Exception ex)
            {
                Host.UI.WriteLine(ConsoleColor.DarkRed, Host.UI.RawUI.BackgroundColor, ex.Message);
            }
        }
        public int LoopSiteCollectionUrls(List<string> sitecollectionUrls, string filePath)
        {
            int counter = 0;
            try
            {                
                string line;
                System.IO.StreamReader file =
                    new System.IO.StreamReader(filePath);
                while ((line = file.ReadLine()) != null)
                {
                    //removes all extra spaces etc. 
                    sitecollectionUrls.Add(line.TrimEnd());
                    counter++;
                }
                file.Close();                
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }
            return counter;
        }


    }
}
