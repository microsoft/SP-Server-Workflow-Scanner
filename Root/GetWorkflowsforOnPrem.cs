using Common;
using Discovery;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Management.Automation;
using System.Security;

namespace Root
{
    public class GetWorkflowsforOnPrem
    {
        public string Url { get; set; }
        public string Scope { get; set; }
        //public string useExistingSummaryFile { get; set; }
        public string userName { get; set; }
        public string password { get; set; }
        public bool OnPrem { get; set; }
        public string DownloadPath { get; set; }
        public bool DownloadForms { get; set; }
        public string DomainName { get; set; }
        public DirectoryInfo analysisFolder;
        public DirectoryInfo downloadedFormsFolder;
        public DirectoryInfo summaryFolder;
        public DataTable dt = new DataTable();

        //public DataTable Execute(PSCredential Credential)
        public DataTable Execute()
        {
            List<string> siteCollectionsUrl = new List<string>();
            try
            {
                Logging.GetInstance().WriteToLogFile(Logging.Info, "Starting to analyze on-premise environment");
                CreateDataTableColumns(dt);
                Console.WriteLine(System.Environment.NewLine);
                Console.WriteLine("Starting to analyze on-premise environment");

                //GetWorkflows();
                if (Scope == "Farm")
                {
                    siteCollectionsUrl = QueryFarm();
                }
                else if (Scope == "WebApplication")
                {
                    siteCollectionsUrl = GetAllWebAppSites();
                }
                else if (Scope == "SiteCollection")
                {
                    siteCollectionsUrl.Add(Url);
                }
                //FindWorkflows(siteCollectionsUrl, Credential);
                FindWorkflows(siteCollectionsUrl);

                Logging.GetInstance().WriteToLogFile(Logging.Info, "***********************************************************************");
                Logging.GetInstance().WriteToLogFile(Logging.Info, "TOTAL WORKFLOWS DISCOVERED : " + dt.Rows.Count.ToString());
                Logging.GetInstance().WriteToLogFile(Logging.Info, "***********************************************************************");

            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);

            }
            return dt;
        }

        public List<string> GetAllWebAppSites()
        {
            List<string> webAppSiteCollectionUrls = new List<string>();
            try
            {
                SPWebApplication objWebApp = null;
                objWebApp = SPWebApplication.Lookup(new Uri(Url));
                if (objWebApp == null)
                {
                    Console.ForegroundColor = ConsoleColor.DarkMagenta;
                    Console.WriteLine("Unable to obtain the object for the Web Application URL provided. Check to make sure the URL provided is correct.");
                    Console.ForegroundColor = ConsoleColor.White;
                    Logging.GetInstance().WriteToLogFile(Logging.Error, "Unable to obtain the object for the Web Application URL provided. SPWebApplication.Lookup(new Uri(Url)) returned NULL");
                }
                else
                {
                    foreach (SPSite site in objWebApp.Sites)
                    {
                        webAppSiteCollectionUrls.Add(site.Url);
                    }
                }


            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }
            return webAppSiteCollectionUrls;
        }

        public List<string> QueryFarm()
        {
            List<string> farmSiteCollectionUrls = new List<string>();
            try
            {
                Logging.GetInstance().WriteToLogFile(Logging.Info, "Starting to query the farm..");
                SPServiceCollection services = SPFarm.Local.Services;
                foreach (SPService curService in services)
                {
                    try
                    {
                        if (curService is SPWebService)
                        {
                            var webService = (SPWebService)curService;
                            if (curService.TypeName.Equals("Microsoft SharePoint Foundation Web Application"))
                            {
                                webService = (SPWebService)curService;
                                SPWebApplicationCollection webApplications = webService.WebApplications;
                                foreach (SPWebApplication webApplication in webApplications)
                                {
                                    // WriteVerbose("Processing WebApplication " + webApplication.DisplayName);
                                    if (webApplication != null)
                                    {
                                        if (false)
                                        {

                                        }
                                        else
                                        {
                                            foreach (SPSite site in webApplication.Sites)
                                            {
                                                try
                                                {
                                                    farmSiteCollectionUrls.Add(site.Url);
                                                    //Console.Write("Gathering InfoPath form locations for Content DB " + contentDatabase.Name + ".....");
                                                    //QueryDatabases(contentDatabase.Name, contentDatabase.DatabaseConnectionString);
                                                    // Console.WriteLine("Done !");
                                                }
                                                catch (Exception ex)
                                                {
                                                    Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                                                    Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                                                    Console.WriteLine("Errored ! See log for details");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                        Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }
            return farmSiteCollectionUrls;
        }
        //public void FindWorkflows(List<string> sitecollectionUrls, PSCredential Credential)
        public void FindWorkflows(List<string> sitecollectionUrls)
        {
            try
            {
                foreach (string url in sitecollectionUrls)
                {
                    ClientContext siteClientContext = null;
                    //siteClientContext = CreateClientContext(url,Credential,DomainName);
                    siteClientContext = CreateClientContext(url, userName,password, DomainName);
                    using (siteClientContext)
                    {
                        bool hasPermissions = false;

                        try

                        {
                            Console.WriteLine(string.Format("Processing: " + url));
                            siteClientContext.ExecuteQueryRetry();
                            hasPermissions = true;
                        }
                        catch (System.Net.WebException webException)
                        {
                            Console.WriteLine(string.Format(webException.Message.ToString() + " on " + url));
                            Logging.GetInstance().WriteToLogFile(Logging.Error, webException.Message.ToString() + " on " + url);
                            Logging.GetInstance().WriteToLogFile(Logging.Error, webException.Message);
                            Logging.GetInstance().WriteToLogFile(Logging.Error, webException.StackTrace);
                        }
                        catch (Microsoft.SharePoint.Client.ClientRequestException clientException)
                        {
                            Console.WriteLine(string.Format(clientException.Message.ToString() + " on " + url));
                            Logging.GetInstance().WriteToLogFile(Logging.Error, clientException.Message.ToString() + " on " + url);
                            Logging.GetInstance().WriteToLogFile(Logging.Error, clientException.Message);
                            Logging.GetInstance().WriteToLogFile(Logging.Error, clientException.StackTrace);
                        }
                        catch (Microsoft.SharePoint.Client.ServerUnauthorizedAccessException unauthorizedException)
                        {
                            Console.WriteLine(string.Format(unauthorizedException.Message.ToString() + " on " + url));
                            Logging.GetInstance().WriteToLogFile(Logging.Error, unauthorizedException.Message.ToString() + " on " + url);
                            Logging.GetInstance().WriteToLogFile(Logging.Error, unauthorizedException.Message);
                            Logging.GetInstance().WriteToLogFile(Logging.Error, unauthorizedException.StackTrace);
                        }

                        if (!hasPermissions)
                            continue;
                        Console.WriteLine(string.Format("Attempting to fetch all the sites and sub sites of  " + url));
                        //WriteVerbose("Trying to get all the sites and subsites of : " + url);
                        IEnumerable<string> expandedSites = siteClientContext.Site.GetAllSubSites();

                        foreach (string site in expandedSites)
                        {
                            //Console.WriteLine(string.Format("Going into " + site));
                            //WriteVerbose("Going into " + site);
                            using (ClientContext ccWeb = siteClientContext.Clone(site))
                            {
                                try
                                {
                                    FindWorkflowPerSite(ccWeb);
                                    //FindInfoPathFormsPerSite(ccWeb);
                                }
                                catch (Microsoft.SharePoint.Client.ServerUnauthorizedAccessException unauthorizedException)
                                {
                                    Logging.GetInstance().WriteToLogFile(Logging.Error, unauthorizedException.Message);
                                    Logging.GetInstance().WriteToLogFile(Logging.Error, unauthorizedException.StackTrace);
                                    Logging.GetInstance().WriteToLogFile(Logging.Error, unauthorizedException.Message.ToString() + " on " + url);
                                    Console.WriteLine(string.Format(unauthorizedException.Message.ToString() + " on " + url));
                                    //WriteWarning(unauthorizedException.Message.ToString() + " on " + url);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }
        }

        public void FindWorkflowPerSite(ClientContext cc)
        {
            try
            {
                var site = cc.Site;
                cc.Load(site);
                cc.ExecuteQueryRetry();

                var web = cc.Web;
                cc.Load(web);
                cc.ExecuteQueryRetry();

                //Host.UI.WriteLine(ConsoleColor.DarkMagenta, Host.UI.RawUI.BackgroundColor, web.Title);
                WorkflowManager.Instance.LoadWorkflowDefaultActions();

                WorkflowDiscovery wfDisc = new WorkflowDiscovery();
                wfDisc.DiscoverWorkflows(cc, dt);
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }
        }

        //internal ClientContext CreateClientContext(string url, string username, SecureString password, string domainName)
        internal ClientContext CreateClientContext(string url, PSCredential Credential, string domainName)
        {
            ClientContext cc = null;
            try
            {
                AuthenticationManager authManager = new AuthenticationManager();
                cc = authManager.GetNetworkCredentialAuthenticatedContext(url, Credential.UserName, Credential.Password, domainName);
                Web web = cc.Web;
                cc.Load(web, website => website.Title);
                cc.ExecuteQuery();
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                return new ClientContext(url)
                {
                };
            }
            return cc;
        }

        internal ClientContext CreateClientContext(string url, string userName, string password, string domainName)
        {
            ClientContext cc = null;
            try
            {
                AuthenticationManager authManager = new AuthenticationManager();
                cc = authManager.GetNetworkCredentialAuthenticatedContext(url, userName, password, domainName);
                Web web = cc.Web;
                cc.Load(web, website => website.Title);
                cc.ExecuteQuery();
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                return new ClientContext(url)
                {
                };
            }
            return cc;
        }

        public void CreateDataTableColumns(DataTable dt)
        {
            try
            {
                dt.Columns.Add("SiteColID");
                dt.Columns.Add("SiteURL");
                dt.Columns.Add("CreatedBy");
                dt.Columns.Add("ModifiedBy");
                dt.Columns.Add("ListTitle");
                dt.Columns.Add("ListUrl");
                dt.Columns.Add("ContentTypeId");
                dt.Columns.Add("ContentTypeName");
                dt.Columns.Add("ItemCount");
                dt.Columns.Add("Scope");
                dt.Columns.Add("Version");
                dt.Columns.Add("WFTemplateName");
                dt.Columns.Add("IsOOBWorkflow");
                dt.Columns.Add("RelativePath");
                dt.Columns.Add("WFID");
                dt.Columns.Add("WebID");
                dt.Columns.Add("WebURL");
                //dt.Columns.Add("ListorLibID");
                //dt.Columns.Add("RelativePath");
                //dt.Columns.Add("IpTemplateName");
                //dt.Columns.Add("InfoPathUsage");
                //dt.Columns.Add("IpID");
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);

            }
        }

        internal static void PopulateAdminAndOwnerColumns(ConcurrentDictionary<string, SiteScanResult> siteScanResults, ConcurrentDictionary<string, WorkflowScanResult> workflowScanResults)
        {
            foreach (var workflowScanResult in workflowScanResults)
            {
                if (siteScanResults.ContainsKey(workflowScanResult.Value.SiteColUrl))
                {
                    var siteScanResult = siteScanResults[workflowScanResult.Value.SiteColUrl];
                    workflowScanResult.Value.Admins = siteScanResult.Admins;
                    workflowScanResult.Value.Owners = siteScanResult.Owners;
                }
            }
        }
    }
}
