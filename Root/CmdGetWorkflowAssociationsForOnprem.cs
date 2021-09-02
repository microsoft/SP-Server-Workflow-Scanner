using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Management.Automation;
using Microsoft.SharePoint.Client;
using System.Security;
using OfficeDevPnP.Core;
using Common;
using Discovery;


namespace Root
{
    public class PwdDynamicParam
    {
        private SecureString pwd;

        [Parameter(Mandatory = true, HelpMessage = "Specify the password as a secure string")]
        public SecureString Password
        {
            get { return pwd; }
            set { pwd = value; }
        }
    }
    /// <summary>
    /// Commandlet to discover the InfoPath forms in 
    /// Onpremise environments that are published to
    /// list and document libraries
    /// </summary>
    /// <returns></returns>   /// 
    ///

    [Cmdlet(VerbsCommon.Get, "WorkflowAssociationsForOnprem")]
    //public class CmdGetWorkflowAssociationsForOnprem : PSCmdlet,IDynamicParameters
    public class CmdGetWorkflowAssociationsForOnprem : PSCmdlet
    {
        private string assessmentScope;
        private static List<string> sitecollectionUrls = new List<string>();

        [Parameter(Mandatory = true, HelpMessage = "Specify the Domain name of the user account")]
        public string DomainName;

        [Parameter(Mandatory = false, HelpMessage = "Specify the password (NOTE : This paramater takes password as plain text. If you like to provide the password as a secure string, ignore this parameter. Youw will be prompted to type in the password as a secure string")]
        public string PasswordPlainText;

        [Parameter(Mandatory = true, HelpMessage = "Specify the user account for authentication")]
        public string UserAccount;


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
        private PwdDynamicParam pwdDyn = null;
        private bool pwdIsPlain = true;
        public DataTable dtWorkflowLocations = new DataTable();


        public object GetDynamicParameters()
        {
            if (Object.Equals(PasswordPlainText, null))
            {
                pwdDyn = new PwdDynamicParam();
                pwdIsPlain = false;
                return pwdDyn;
            }
            else
                return null;

        }
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
                    if (String.IsNullOrEmpty(SiteCollectionUrl))
                    {
                        assessmentScope = "Farm";
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

                    //Logging.GetInstance().WriteToLogFile(Logging.Info, "Beginning assessment..");
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
                    objonPrem.DownloadPath = AssessmentOutputFolder;
                    //Set UserName & Password
                    objonPrem.userName = UserAccount;
                    if (pwdIsPlain)
                        objonPrem.password = PasswordPlainText;
                    else
                        objonPrem.password = pwdDyn.Password.ToString();
                    //dtWorkflowLocations = objonPrem.Execute(Credential);
                    dtWorkflowLocations = objonPrem.Execute();
                    //Save the CSV file
                    string csvFilePath = string.Concat(AssessmentOutputFolder, ops.summaryFolder, ops.summaryFile);
                    ops.WriteToCsvFile(dtWorkflowLocations, csvFilePath);
                }
                else if (userInput.Equals("n"))
                {
                    Host.UI.WriteLine(ConsoleColor.Cyan, Host.UI.RawUI.BackgroundColor, "Operation aborted as per your input !");
                }              
                //New Code Ends


                //Create Assessment folders
                /* old code commented out
                ops.CreateDirectoryStructure(AssessmentOutputFolder);

                ops.CreateDataTableColumns(dt);
                string csvFilePath = string.Concat(AssessmentOutputFolder, ops.summaryFolder, ops.summaryFile);
                string downloadXomlFolderPath = string.Concat(AssessmentOutputFolder, ops.downloadedFormsFolder);
                sitecollectionUrls.Clear();
                int scCount = LoopSiteCollectionUrls(sitecollectionUrls, SiteCollectionURLFilePath);
                if (scCount == 0)
                {
                    WriteWarning(string.Format("The text file located at {0} does not contain Site Collection URLs and appears to be empty !", SiteCollectionURLFilePath));
                }
                else
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    ClientContext cc = null;

                    foreach (var scUrl in sitecollectionUrls)
                    {
                        if (pwdIsPlain)
                            cc = authManager.GetNetworkCredentialAuthenticatedContext(scUrl, UserAccount, PasswordPlainText, DomainName);
                        else
                            cc = authManager.GetNetworkCredentialAuthenticatedContext(scUrl, UserAccount, pwdDyn.Password, DomainName);

                        Web web = cc.Web;
                        cc.Load(web, website => website.Title);
                        cc.ExecuteQuery();

                        Host.UI.WriteLine(ConsoleColor.DarkMagenta, Host.UI.RawUI.BackgroundColor, web.Title);
                        WorkflowManager.Instance.LoadWorkflowDefaultActions();

                        WorkflowDiscovery wfDisc = new WorkflowDiscovery();
                        wfDisc.DiscoverWorkflows(cc, dt);
                    }
                    //Save the CSV file
                    ops.WriteToCsvFile(dt, csvFilePath);
                }
                */

            }
            catch (Exception ex)
            {
                Host.UI.WriteLine(ConsoleColor.DarkRed, Host.UI.RawUI.BackgroundColor, ex.Message);
            }
        }
        private static SecureString GetPassword()
        {
            ConsoleKeyInfo info;
            //Get the user's password as a SecureString    
            SecureString securePassword = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return securePassword;
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
                    //System.Console.WriteLine(line);
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
