using Common;
using Discovery;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using File = Microsoft.SharePoint.Client.File;

namespace Root
{
    public class GetWorkflowForSPOnline
    {

        /// <summary>
        /// Open the site collection file and store in a collection variable
        /// Read the file and display it line by line.  
        /// </summary>
        /// <param name="sitecollectionUrls"></param>
        public void ReadInfoPathOnlineSiteCollection(List<string> sitecollectionUrls, string filePath)
        {
            try
            {
                int counter = 0;
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
        }
        private static readonly string FormBaseContentType = "0x010101";
        public DataTable dt = new DataTable();

        /// <summary>
        /// Discover all sites present in the tenant and save them into a collection
        /// </summary>
        /// <param name="TenantName"></param>
        /// <param name="Credential"></param>
        /// <returns></returns>
        public List<string> GetAllSPOTenantSites(string TenantName, PSCredential Credential)
        {
            List<string> sites = new List<string>();
            try
            {
                string tenantAdminUrl = "https://" + TenantName + "-admin.sharepoint.com/";
                ClientContext ctx = null;
                ctx = CreateClientContext(tenantAdminUrl, Credential.UserName, Credential.Password);
                var site = ctx.Site;
                ctx.Load(site);
                ctx.ExecuteQueryRetry();
                var web = site.RootWeb;
                var list = web.Lists.GetByTitle("DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS");
                ctx.Load(list);
                ctx.ExecuteQuery();
                //Console.WriteLine("List name :{0}", list.Title);
                var camlQuery = new CamlQuery();
                var items = list.GetItems(camlQuery);
                ctx.Load(items);
                ctx.ExecuteQuery();
                foreach (var item in items)
                {
                    //Console.WriteLine("Site Url {0}", item["SiteUrl"]);
                    if (item["SiteUrl"] != null)
                        sites.Add(item["SiteUrl"].ToString());
                }
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                //Console.WriteLine(ex.Message);
            }
            return sites;

        }

        public List<string> GetAllTenantSites(string TenantName, PSCredential Credential)
        {
            List<string> sites = new List<string>();
            try
            {
                string tenantAdminUrl = "https://" + TenantName + "-admin.sharepoint.com/";
                ClientContext ctx = null;
                ctx = CreateClientContext(tenantAdminUrl, Credential.UserName, Credential.Password);
                Tenant tenant = new Tenant(ctx);
                SPOSitePropertiesEnumerable siteProps = tenant.GetSitePropertiesFromSharePoint("0", true);
                ctx.Load(siteProps);
                ctx.ExecuteQuery();
                int count = 0;
                foreach (var site in siteProps)
                {
                    sites.Add(site.Url);
                    count++;
                }
                Console.WriteLine("Total Site {0}", count);
                return sites;
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);

            }
            return sites;
        }
        /// <summary>
        /// SPO 
        /// </summary>
        public string DownloadPath { get; set; }
        public bool DownloadForms { get; set; }
        public DirectoryInfo analysisFolder;
        public DirectoryInfo downloadedFormsFolder;
        public DirectoryInfo summaryFolder;

        public void FindWorkflows(List<string> sitecollectionUrls, PSCredential Credential)
        {
            try
            {
                foreach (string url in sitecollectionUrls)
                {
                    ClientContext siteClientContext = null;
                    siteClientContext = CreateClientContext(url, Credential.UserName, Credential.Password);
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
        /// <summary>
        /// 
        /// </summary>
        /// <param name="url"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        internal ClientContext CreateClientContext(string url, string username, SecureString password)
        {
            try
            {
                var credentials = new SharePointOnlineCredentials(
                                       username,
                                       password);

                return new ClientContext(url)
                {
                    Credentials = credentials
                };
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                return new ClientContext(url)
                {
                };
            }


        }

        internal DataTable FindInfoPathFormsPerSite(ClientContext cc)
        {
            int counter = 0;
            //DataTable dt = new DataTable();            
            try
            {
                var site = cc.Site;
                cc.Load(site);
                cc.ExecuteQueryRetry();

                var web = cc.Web;
                cc.Load(web);
                cc.ExecuteQueryRetry();


                var lists = cc.Web.GetListsToScan(showHidden: true);
                Logging.GetInstance().WriteToLogFile(Logging.Info, string.Format("Started Scanning site {0}", site.Url));

                foreach (var list in lists)
                {
                    try
                    {
                        cc.Load(list);
                        cc.ExecuteQueryRetry();
                        if (list.BaseTemplate == (int)ListTemplateType.XMLForm ||
                            (!string.IsNullOrEmpty(list.DocumentTemplateUrl) && list.DocumentTemplateUrl.EndsWith(".xsn", StringComparison.InvariantCultureIgnoreCase))
                           )
                        {
                            try
                            {
                                DataRow row = dt.NewRow();

                                File file = GetFileId(cc, list.DocumentTemplateUrl);
                                if (file != null)
                                {
                                    row["IpID"] = file.UniqueId;
                                }
                                // Form libraries depend on InfoPath

                                row["SiteColID"] = site.Id;
                                row["SiteURL"] = site.Url;
                                row["WebID"] = web.Id;
                                row["WebURL"] = web.Url;
                                row["ListorLibID"] = list.Id;
                                row["RelativePath"] = list.RootFolder.ServerRelativeUrl + "/Forms";
                                row["IpTemplateName"] = Path.GetFileName(list.DocumentTemplateUrl);
                                row["InfoPathUsage"] = "FormLibrary";
                                //row["FileName"] = list.Properties.FieldValues["_ipfs_solutionName"].ToString();
                                row["ItemCount"] = list.ItemCount;
                                dt.Rows.Add(row);
                                Logging.GetInstance().WriteToLogFile(Logging.Info, string.Format("Found InfoPath Forms Library {0}", list.DefaultViewUrl));
                                counter++;
                            }
                            catch (Exception ex)
                            {
                                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message.ToString() + " on " + web.Url);
                                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);

                            }
                        }
                        #region contentType Code
                        /*
                        else if (list.BaseTemplate == (int)ListTemplateType.DocumentLibrary || list.BaseTemplate == (int)ListTemplateType.WebPageLibrary)
                        {
                            try
                            {                                
                                // verify if a form content type was attached to this list
                                cc.Load(list, p => p.ContentTypes.Include(c => c.Id, c => c.DocumentTemplateUrl));
                                //cc.Load(list);
                                cc.ExecuteQueryRetry();

                                var formContentTypeFound = list.ContentTypes.Where(c => c.Id.StringValue.StartsWith(FormBaseContentType, StringComparison.InvariantCultureIgnoreCase)).OrderBy(c => c.Id.StringValue.Length).FirstOrDefault();
                                if (formContentTypeFound != null)
                                {
                                    // Form libraries depend on InfoPath
                                    DataRow row = dt.NewRow();
                                    if (formContentTypeFound.DocumentTemplateUrl.EndsWith(".xsn"))
                                    {
                                        File file = GetFileId(cc, list, formContentTypeFound.DocumentTemplateUrl);
                                        if (file != null)
                                        {
                                            row["IpID"] = file.UniqueId;
                                        }
                                    }
                                    row["SiteColID"] = site.Id;
                                    row["SiteURL"] = site.Url;
                                    row["WebID"] = web.Id;
                                    row["WebURL"] = web.Url;
                                    row["ListorLibID"] = list.Id;
                                    row["RelativePath"] = list.RootFolder.ServerRelativeUrl;
                                    row["InfoPathUsage"] = "ContentType";
                                    row["IpTemplateName"] = Path.GetFileName(formContentTypeFound.DocumentTemplateUrl);
                                    //row["FileName"] = list.Properties.FieldValues["_ipfs_solutionName"].ToString();
                                    row["ItemCount"] = list.ItemCount;
                                    dt.Rows.Add(row);
                                    Logging.GetInstance().WriteToLogFile(Logging.Info, string.Format("Found Content Type {0}", formContentTypeFound.DocumentTemplateUrl));
                                    counter++;
                                }

                            }
                            catch (Exception ex)
                            {
                                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                            }
                        }
                        */
                        #endregion
                        else if (list.BaseTemplate == (int)ListTemplateType.GenericList)
                        {
                            try
                            {
                                Folder folder = cc.Web.GetFolderByServerRelativeUrl($"{list.RootFolder.ServerRelativeUrl}/Item");
                                cc.Load(folder, p => p.Properties);
                                cc.ExecuteQueryRetry();

                                if (folder.Properties.FieldValues.ContainsKey("_ipfs_infopathenabled") && folder.Properties.FieldValues.ContainsKey("_ipfs_solutionName"))
                                {
                                    bool infoPathEnabled = true;
                                    if (bool.TryParse(folder.Properties.FieldValues["_ipfs_infopathenabled"].ToString(), out bool infoPathEnabledParsed))
                                    {
                                        infoPathEnabled = infoPathEnabledParsed;
                                    }
                                    // Form libraries depend on InfoPath
                                    if (infoPathEnabled)
                                    {
                                        DataRow row = dt.NewRow();
                                        string templateUrl = list.RootFolder.ServerRelativeUrl + "/Item/" + folder.Properties["_ipfs_solutionName"];
                                        File file = GetFileId(cc, templateUrl);
                                        if (file != null)
                                        {
                                            row["IpID"] = file.UniqueId;
                                        }
                                        row["SiteColID"] = site.Id;
                                        row["SiteURL"] = site.Url;
                                        row["WebID"] = web.Id;
                                        row["WebURL"] = web.Url;
                                        row["ListorLibID"] = list.Id;
                                        row["InfoPathUsage"] = "GenericList";
                                        row["RelativePath"] = list.RootFolder.ServerRelativeUrl + "/Item";
                                        row["IpTemplateName"] = folder.Properties["_ipfs_solutionName"];
                                        row["ItemCount"] = list.ItemCount;
                                        dt.Rows.Add(row);
                                        Logging.GetInstance().WriteToLogFile(Logging.Info, string.Format("Found Generic List InfoPath form {0}", list.DefaultViewUrl));
                                        counter++;

                                    }
                                }
                            }
                            catch (ServerException ex)
                            {
                                if (((ServerException)ex).ServerErrorTypeName == "System.IO.FileNotFoundException")
                                {
                                    // Ignore
                                }
                                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message.ToString() + " on " + web.Url);
                                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message.ToString() + " on " + web.Url);
                        Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                        Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                    }
                }
                Logging.GetInstance().WriteToLogFile(Logging.Info, string.Format("Found a total of {0} InfoPath forms in site {1}", counter, web.Url));
                Logging.GetInstance().WriteToLogFile(Logging.Info, string.Format("Completed Scanning site {0}", site.Url));
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);

            }
            //counter = 0;
            return dt;
        }

        private static File GetFileId(ClientContext cc, string documentTemplateUrl)
        {
            //Guid fileId = new Guid();
            File file = null;
            try
            {
                var spfileLocation = String.Concat(documentTemplateUrl);

                file = cc.Web.GetFileByServerRelativeUrl(spfileLocation);
                // Then getting the file using the server-relative Url of the web object
                cc.Load(file);
                cc.ExecuteQuery();

            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message.ToString() + " on " + documentTemplateUrl);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }
            return file;
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

        public DataTable Execute(PSCredential Credential, List<string> sitecollectionUrls)
        {
            try
            {
                Logging.GetInstance().WriteToLogFile(Logging.Info, "Starting to analyze SharePoint Online environment");

                Console.WriteLine(System.Environment.NewLine);
                Console.WriteLine("Starting to analyze SharePoint Online environment");
                CreateDataTableColumns(dt);
                //FindInfoPathForms(sitecollectionUrls, Credential);
                FindWorkflows(sitecollectionUrls, Credential);                    
                Logging.GetInstance().WriteToLogFile(Logging.Info, "***********************************************************************");
                Logging.GetInstance().WriteToLogFile(Logging.Info, "TOTAL WORKFLOWS DISCOVERED : " + dt.Rows.Count.ToString());
                Logging.GetInstance().WriteToLogFile(Logging.Info, "***********************************************************************");

                Console.WriteLine(System.Environment.NewLine);
                Console.WriteLine("***********************************************************************");
                Console.WriteLine("TOTAL WORKFLOWS DISCOVERED : " + dt.Rows.Count.ToString());
                Console.WriteLine("***********************************************************************");
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }
            return dt;
        }

        public void DownloadInfoPathForms(DataTable infopathFormLocations, PSCredential Credential)
        {
            try
            {
                Console.WriteLine(System.Environment.NewLine);
                Console.WriteLine("***********************************************************************");
                Console.WriteLine("Beginning to download the InfoPath template forms locally.. ");
                Console.WriteLine("***********************************************************************");
                Console.WriteLine(System.Environment.NewLine);
                Console.WriteLine("Downloading forms...please wait.. ");
                Console.WriteLine(System.Environment.NewLine);



                Logging.GetInstance().WriteToLogFile(Logging.Info, "***********************************************************************");
                Logging.GetInstance().WriteToLogFile(Logging.Info, "Beginning to download the InfoPath template forms locally.. ");
                Logging.GetInstance().WriteToLogFile(Logging.Info, "***********************************************************************");

                int formsDownloadCounter = 0;
                string spfileLocation = string.Empty;
                Operations ops = new Operations();
                foreach (DataRow infoPathFormLocation in infopathFormLocations.Rows)
                {
                    try
                    {
                        spfileLocation = String.Concat(infoPathFormLocation["RelativePath"].ToString(), "/", infoPathFormLocation["IpTemplateName"].ToString());

                        ClientContext context = null;
                        context = CreateClientContext(infoPathFormLocation["WebURL"].ToString(), Credential.UserName, Credential.Password);
                        using (context)
                        {
                            List list = context.Web.GetListById(Guid.Parse(infoPathFormLocation["ListorLibID"].ToString()));
                            context.Load(list);
                            context.ExecuteQueryRetry();

                            var file = GetFileId(context, spfileLocation);
                            //var file = context.Web.GetFileByServerRelativeUrl(infoPathFormLocation["RelativePath"].ToString());
                            //context.Load(file);
                            //context.ExecuteQuery();

                            string folderName = infoPathFormLocation["SiteColID"].ToString();
                            //var dirInfo = System.IO.Directory.CreateDirectory(string.Concat(DownloadPath, @"\DownloadedForms", @"\", infoPathFormLocation["SiteColID"].ToString()));
                            var dirInfo = System.IO.Directory.CreateDirectory(string.Concat(DownloadPath, ops.downloadedFormsFolder, @"\", infoPathFormLocation["SiteColID"].ToString()));
                            string filename = file.Name;

                            string filePath = string.Concat(DownloadPath, @"\DownloadedForms", @"\", infoPathFormLocation["SiteColID"].ToString(), @"\", infoPathFormLocation["IpID"].ToString(), "_", filename);

                            //var dirInfo = System.IO.Directory.CreateDirectory(string.Concat(DownloadPath, @"\", infoPathFormLocation["SiteID"], @"\", infoPathFormLocation["WebID"].ToString()));
                            //string localfile = string.Concat(DownloadPath, @"\", infoPathFormLocation["SiteID"].ToString(), @"\", infoPathFormLocation["WebID"].ToString(), @"\", infoPathFormLocation["ListorLibID"].ToString(), "_", filename);
                            Logging.GetInstance().WriteToLogFile(Logging.Info, "Starting to download InfoPath located at " + list.DefaultViewUrl);
                            //Logging.GetInstance().WriteToLogFile(Logging.Info, "Downloading the form locallay to the location " + filePath);
                            try
                            {
                                //Downloading the file to the specified location
                                using (var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, spfileLocation))
                                {
                                    using (FileStream writeStream = System.IO.File.Open(filePath, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write, System.IO.FileShare.ReadWrite))
                                    {
                                        fileInfo.Stream.CopyTo(writeStream);
                                        formsDownloadCounter++;
                                    }
                                }
                                Logging.GetInstance().WriteToLogFile(Logging.Info, "Finished to download InfoPath located at " + list.DefaultViewUrl);
                                //Logging.GetInstance().WriteToLogFile(Logging.Info, "Finished to download InfoPath located at " + web.Url + "/" + list.DefaultViewUrl);
                                //Logging.GetInstance().WriteToLogFile(Logging.Info, "Download complete for the form located in " + spfileLocation);

                            }
                            catch (Exception exception)
                            {
                                if (infoPathFormLocation["WebURL"] != null)
                                {
                                    Logging.GetInstance().WriteToLogFile(Logging.Error, exception.Message + " on " + infoPathFormLocation["WebURL"].ToString());
                                }
                                Logging.GetInstance().WriteToLogFile(Logging.Error, exception.Message);
                                Logging.GetInstance().WriteToLogFile(Logging.Error, exception.StackTrace);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (infoPathFormLocation["WebURL"] != null)
                        {
                            Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message + " on " + infoPathFormLocation["WebURL"].ToString());
                        }
                        Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                        Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                    }
                }
                Console.WriteLine("Downloading forms complete !");
                Console.WriteLine(System.Environment.NewLine);
                Console.WriteLine("***********************************************************************");
                Console.WriteLine("INFOPATH FORMS DOWNLOADED : " + formsDownloadCounter);
                Console.WriteLine("***********************************************************************");
                Console.WriteLine(System.Environment.NewLine);

                Logging.GetInstance().WriteToLogFile(Logging.Info, "***********************************************************************");
                Logging.GetInstance().WriteToLogFile(Logging.Info, "INFOPATH FORMS DOWNLOADED : " + formsDownloadCounter);
                Logging.GetInstance().WriteToLogFile(Logging.Info, "***********************************************************************");
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);

            }
            //Logging.GetInstance().WriteToLogFile(Logging.Info, "STARTING TO DOWNLOAD INFOPATH FORMS. TOTAL COUNT: " + infopathFormLocations.Rows.Count);

            //Load the Summary CSV to InfoPathFormLocations Collection

            //foreach (var location in infopathFormLocations)

        }
    }
}
