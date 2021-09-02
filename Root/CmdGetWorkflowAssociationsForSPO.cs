using Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace Root
{
    [Cmdlet(VerbsCommon.Get, "WorkflowAssociationsForSPO")]
    public class CmdGetWorkflowAssociationsForSPO : PSCmdlet
    {
        private static List<string> sitecollectionUrls = new List<string>();
        //[Parameter(Mandatory = true, HelpMessage = "Credentials to login to SPO for InfoPath Online Assessment")]
        ////[System.Management.Automation.PSCredential]
        //[System.Management.Automation.CredentialAttribute()]
        //PSCredential Credential;
        private string assessmentScope;
        [Parameter(Mandatory = true, ParameterSetName = "Credential")]
        public PSCredential Credential;

        [Parameter(Mandatory = false, HelpMessage = "Specify the tenant name of your SPO site collections")]
        public string TenantName;

        [Parameter(Mandatory = false, HelpMessage = "Specify the file path of a text file containing all SPO site collection URLs")]
        public string SiteCollectionURLFilePath;

        [Parameter(Mandatory = true, HelpMessage = "Set this switch to $true if you would like to download the Workflows for detailed assessment")]
        public bool DownloadWorkflows
        {
            get { return downloadWorkflows; }
            set { downloadWorkflows = value; }
        }
        private bool downloadWorkflows;

        [Parameter(Mandatory = true, HelpMessage = @"The path where the Assessment Summary, logs, Workflows forms are downloaded (if DownloadWorkflow parameter is set to true) for analyzing (e.g. F:\temp\Workflows")]
        public string AssessmentOutputFolder;
        private string logFolderPath;
        public DirectoryInfo logFolder;

        protected override void BeginProcessing()
        {
            //Create required folders
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
                try
                {

                    if (String.IsNullOrEmpty(TenantName) && String.IsNullOrEmpty(SiteCollectionURLFilePath))
                    {
                        WriteWarning("Provide either the Tenant Name or the Site Collection File containing SPO URLs !");
                    }
                    else if (!String.IsNullOrEmpty(TenantName))
                    {
                        assessmentScope = "Tenant";
                        BeginToAssess();
                    }
                    else
                    {
                        assessmentScope = "SiteCollection";
                        BeginToAssess();
                    }

                }
                catch (Exception ex)
                {
                    Host.UI.WriteLine(ConsoleColor.DarkRed, Host.UI.RawUI.BackgroundColor, ex.Message);
                }

            }
            catch (Exception ex)
            { }
        }

        protected void BeginToAssess()
        {
            GetWorkflowForSPOnline objSPOnline = new GetWorkflowForSPOnline();
            Operations ops = new Operations();
            try
            {
                string downloadXomlFolderPath = string.Concat(AssessmentOutputFolder, ops.downloadedFormsFolder);
                //Scope level code getting addded
                string userInput = string.Empty;
                Console.WriteLine(System.Environment.NewLine);
                Collection<PSObject> op = new Collection<PSObject>();
                List<string> allSPOTenantSites = new List<string>();

                if (assessmentScope == "Tenant")
                {
                    allSPOTenantSites = ops.GetAllTenantSites(TenantName, Credential);
                    Host.UI.WriteLine(ConsoleColor.Cyan, Host.UI.RawUI.BackgroundColor, "The assessment is scoped to run at " + assessmentScope +
" level. There are altogether " + allSPOTenantSites.Count + " sites in your tenant. Would you like to proceed? [Y] to continue, [N] to abort.");
                    op = this.InvokeCommand.InvokeScript("Read-Host");
                    userInput = op[0].ToString().ToLower();
                }
                else
                {
                    Host.UI.WriteLine(ConsoleColor.Cyan, Host.UI.RawUI.BackgroundColor, "The assessment is scoped to run at " + assessmentScope +
" level. Would you like to proceed? [Y] to continue, [N] to abort.");
                    op = this.InvokeCommand.InvokeScript("Read-Host");
                    userInput = op[0].ToString().ToLower();
                }
                while (!userInput.Equals("y") && !userInput.Equals("n"))
                {
                    Host.UI.WriteLine(ConsoleColor.Cyan, Host.UI.RawUI.BackgroundColor, "Invalid input. Press [Y] to continue, [N] to abort.");
                    op = this.InvokeCommand.InvokeScript("Read-Host");
                    userInput = op[0].ToString().ToLower();
                }
                if (userInput.Equals("y"))
                {
                    DataTable dtWorkflowLocations = new DataTable();
                    //Clear the collection before running
                    sitecollectionUrls.Clear();
                    //Create Assessment output folders
                    ops.CreateDirectoryStructure(AssessmentOutputFolder);
                    if (assessmentScope == "Tenant")
                    {
                        sitecollectionUrls = allSPOTenantSites;
                    }
                    else
                    {
                        objSPOnline.ReadInfoPathOnlineSiteCollection(sitecollectionUrls, SiteCollectionURLFilePath);
                    }
                    if (sitecollectionUrls.Count == 0)
                    {
                        Host.UI.WriteLine(ConsoleColor.Red, Host.UI.RawUI.BackgroundColor,
                            string.Format("Site Collection URLs at {0} text file were empty, please update the text file with SPO Workflows site URL.", SiteCollectionURLFilePath));
                        Logging.GetInstance().WriteToLogFile(Logging.Info, string.Format("Site Collection URLs at {0} text file were empty, please update the text file with SPO Workflows site URL.", SiteCollectionURLFilePath));
                    }
                    else
                    {
                        string csvFilePath = string.Concat(AssessmentOutputFolder, ops.summaryFolder, ops.summaryFile);
                        if (DownloadWorkflows)
                        { }
                        else
                        {

                            objSPOnline.DownloadPath = AssessmentOutputFolder;
                            objSPOnline.DownloadForms = DownloadWorkflows;
                            dtWorkflowLocations = objSPOnline.Execute(Credential, sitecollectionUrls);
                            //Save the CSV file
                            ops.WriteToCsvFile(dtWorkflowLocations, csvFilePath);
                        }
                    }
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
    }
}
