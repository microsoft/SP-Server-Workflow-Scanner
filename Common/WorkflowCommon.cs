//using Microsoft.Deployment.Compression.Cab;
//using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Data;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Management.Automation;
using Microsoft.Online.SharePoint.TenantAdministration;

namespace Common
{
    
    public class Operations

    {
        /// <summary>
        /// SPO Properties
        /// </summary>
        public string DownloadPath { get; set; }
        public bool DownloadForms { get; set; }
        public string summaryFile = @"\WorkflowDiscovery.csv";
        public string logFolder = @"\Logs";
        public string downloadedFormsFolder = @"\DownloadedWorkflows";
        public string analysisFolder = @"\Analysis";
        public string summaryFolder = @"\Summary";
        public string analysisOutputFile = @"\WorkflowComparisonDetails.csv";
        public string compOutputFile = @"\WorkflowComparison.csv";
        public DataTable dt = new DataTable();


        public void SaveXamlFile(string xamlContent, Web web, string wfName, string scope, string folderPath)
        {
            try
            {
                string fileName = web.Id + "-" + wfName + "-" + scope+".xoml";
                string filePath = folderPath + "\\"+ fileName;
                System.IO.File.WriteAllText(filePath, xamlContent);
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }        
        }
        /// <summary>
        /// Create Data Table
        /// </summary>
        /// <param name="dt"></param>
        public void CreateDataTableColumns(DataTable dt)
        {
            try
            {
                dt.Columns.Add("SiteColID");
                dt.Columns.Add("SiteURL");
                dt.Columns.Add("ListTitle");
                dt.Columns.Add("ListUrl");
                dt.Columns.Add("ContentTypeId");
                dt.Columns.Add("ContentTypeName");
                dt.Columns.Add("Scope");
                dt.Columns.Add("Version");
                dt.Columns.Add("WFTemplateName");
                dt.Columns.Add("WorkFlowName");
                dt.Columns.Add("IsOOBWorkflow");
          //    dt.Columns.Add("RelativePath");
                dt.Columns.Add("WFID");
                dt.Columns.Add("WebID");
                dt.Columns.Add("WebURL");
                dt.Columns.Add("Enabled");
                dt.Columns.Add("HasSubscriptions");
                dt.Columns.Add("ConsiderUpgradingToFlow");
                dt.Columns.Add("ToFLowMappingPercentage");
                dt.Columns.Add("UsedActions");
                dt.Columns.Add("ActionCount");
                dt.Columns.Add("AllowManual");
                dt.Columns.Add("AutoStartChange");
                dt.Columns.Add("AutoStartCreate");
                dt.Columns.Add("LastDefinitionModifiedDate");
                dt.Columns.Add("LastSubsrciptionModifiedDate");
                dt.Columns.Add("AssociationData");
                //dt.Columns.Add("Complexity");
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

        /// <summary>
        /// Create Data Table
        /// </summary>
        /// <param name="dt"></param>
        public void AddRowToDataTable(WorkflowScanResult workflowScanResult, DataTable dt, string version, string scope, string wfName, string wfID, bool IsOOBWF, Web web)
        {
            DataRow dr = dt.NewRow();
            try
            {
                if (workflowScanResult.SiteColUrl == null && web.Url != null)
                {
                    dr["SiteURL"] = web.Url;
                    dr["SiteColID"] = web.Id;
                }
                else
                dr["SiteURL"] = workflowScanResult.SiteColUrl;
                //dr["WebURL"] = workflowScanResult.SiteURL;
            //    dr["CreatedBy"] = workflowScanResult.CreatedBy;
            //    dr["ModifiedBy"] = workflowScanResult.ModifiedBy;
                dr["WebURL"] = web.Url;
                dr["ListTitle"] = workflowScanResult.ListTitle;
                dr["ListUrl"] = workflowScanResult.ListUrl;
                dr["ContentTypeId"] = workflowScanResult.ContentTypeId;
                dr["ContentTypeName"] = workflowScanResult.ContentTypeName;
                dr["Scope"] = scope;
                dr["Version"] = version;
                dr["WFTemplateName"] = wfName;
                dr["WorkFlowName"] = workflowScanResult.SubscriptionName;
                dr["IsOOBWorkflow"] = IsOOBWF;
                dr["Enabled"] = workflowScanResult.Enabled;   // adding for is enabled 
                dr["WFID"] = wfID;
                dr["WebID"] = web.Id;
                dr["HasSubscriptions"] = workflowScanResult.HasSubscriptions;   // adding for subscriptions 
                string sUsedActions = "";
                // AM need to refactor into a helper function
                if (workflowScanResult.UsedActions != null)
                {
                    foreach (var item in workflowScanResult.UsedActions)
                    {
                        sUsedActions = item.ToString()+";"+ sUsedActions;
                    }
                }
                dr["ToFLowMappingPercentage"] = workflowScanResult.ToFLowMappingPercentage;   // adding for percentange upgradable to flow 
                dr["ConsiderUpgradingToFlow"] = workflowScanResult.ConsiderUpgradingToFlow;   // adding for consider upgrading to flow 
                dr["UsedActions"] = sUsedActions;   // adding for UsedActions
                dr["ActionCount"] = workflowScanResult.ActionCount;   // adding for ActionCount
                dr["AllowManual"] = workflowScanResult.AllowManual;
                dr["AutoStartChange"] = workflowScanResult.AutoStartChange;
                dr["AutoStartCreate"] = workflowScanResult.AutoStartCreate;
                dr["LastDefinitionModifiedDate"] = workflowScanResult.LastDefinitionEdit;
                dr["LastSubsrciptionModifiedDate"] = workflowScanResult.LastSubscriptionEdit;
                dr["AssociationData"] = workflowScanResult.AssociationData;



                //dr["Complexity"] = "High";   // adding placeholder for Complexity
                dt.Rows.Add(dr);
                
                //dt.Columns.Add("SiteColID");
                //dt.Columns.Add("SiteURL");
                //dt.Columns.Add("ListTitle");
                //dt.Columns.Add("ListUrl");
                //dt.Columns.Add("ContentTypeId");
                //dt.Columns.Add("ContentTypeName");
                //dt.Columns.Add("ItemCount");
                //dt.Columns.Add("Scope");
                //dt.Columns.Add("Version");
                //dt.Columns.Add("WFTemplateName");
                //dt.Columns.Add("IsOOBWorkflow");
                //dt.Columns.Add("RelativePath");
                //dt.Columns.Add("WFID");
                //dt.Columns.Add("WebID");
                //dt.Columns.Add("WebURL");
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);

            }
        }

        /// <summary>
        /// Creates 3 levels of folder at the downloaded path provided by the user
        /// 1. Analysis
        /// 2. DownloadedForms
        /// 3. Summary
        /// </summary>
        /// <param name="downloadPath"></param>
        public void CreateDirectoryStructure(string downloadPath)
        {
            try
            {
                if (!Directory.Exists(string.Concat(downloadPath, analysisFolder)))
                {
                    DirectoryInfo analysisFolder1 = System.IO.Directory.CreateDirectory(string.Concat(downloadPath, analysisFolder));
                    Logging.GetInstance().WriteToLogFile(Logging.Info, "Analysis folder created");
                }
                if (!Directory.Exists(string.Concat(downloadPath, downloadedFormsFolder)))
                {
                    DirectoryInfo downloadedFormsFolder1 = System.IO.Directory.CreateDirectory(string.Concat(downloadPath, downloadedFormsFolder));
                    Logging.GetInstance().WriteToLogFile(Logging.Info, "DownloadedForms folder created");

                }
                if (!Directory.Exists(string.Concat(downloadPath, summaryFolder)))
                {
                    DirectoryInfo summaryFolder1 = System.IO.Directory.CreateDirectory(string.Concat(downloadPath, summaryFolder));
                    Logging.GetInstance().WriteToLogFile(Logging.Info, "Summary folder created");

                }
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }
        }

        public DataTable ConvertCSVToDataTable(string csvLocation)
        {
            DataTable dt = new DataTable();
            try
            {
                var Lines = System.IO.File.ReadAllLines(csvLocation);
                string[] Fields;
                Fields = Lines[0].Split(new char[] { ',' });
                int Cols = Fields.GetLength(0);
                for (int i = 0; i < Cols; i++)
                    dt.Columns.Add(Fields[i].ToLower(), typeof(string));
                DataRow Row;
                for (int i = 1; i < Lines.GetLength(0); i++)
                {
                    Fields = Lines[i].Split(new char[] { ',' });
                    Row = dt.NewRow();
                    for (int f = 0; f < Cols; f++)
                        Row[f] = Fields[f].Split('"')[1];
                    dt.Rows.Add(Row);
                }
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }
            return dt;
        }


        /// <summary>
        /// The DataSet returned from the content database is stored in a datatable
        /// The datatable is then saved into a CSV file that gets stored in the Summary folder
        /// that gets created at the location of download path supplied by the users
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="filePath"></param>
        public void WriteToCsvFile(DataTable dataTable, string filePath)
        {
            try
            {
                Logging.GetInstance().WriteToLogFile(Logging.Info, "Preparing to create the CSV at " + filePath);
                StringBuilder fileContent = new StringBuilder();

                foreach (var col in dataTable.Columns)
                {
                    fileContent.Append(col.ToString() + ",");
                }

                fileContent.Replace(",", System.Environment.NewLine, fileContent.Length - 1, 1);

                foreach (DataRow dr in dataTable.Rows)
                {
                    foreach (var column in dr.ItemArray)
                    {
                        fileContent.Append("\"" + column.ToString() + "\",");
                    }

                    fileContent.Replace(",", System.Environment.NewLine, fileContent.Length - 1, 1);
                }

                System.IO.File.WriteAllText(filePath, fileContent.ToString());
                Logging.GetInstance().WriteToLogFile(Logging.Info, string.Format("CSV File created at {0}", filePath));
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
        /// <param name="s"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public void WriteProgress(string s, int x, int y)
        {
            int origRow = Console.CursorTop;
            int origCol = Console.CursorLeft;
            // Console.WindowWidth = 10;  // this works. 
            int width = Console.WindowWidth;
            //x = x % width;
            try
            {
                Console.SetCursorPosition(x, y);
                //Console.SetCursorPosition(origCol, origRow);
                Console.Write(s);
            }
            catch (ArgumentOutOfRangeException e)
            {

            }
            finally
            {
                try
                {
                    Console.SetCursorPosition(origRow, origCol);
                }
                catch (ArgumentOutOfRangeException e)
                {
                }
            }
        }

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
    }

 


    /*
    public class InfoPathScanResult
    {
        public string SiteColUrl { get; set; }
        public string SiteURL { get; set; }
        public string InfoPathUsage { get; set; }
        public string ListTitle { get; set; }
        public string ListId { get; set; }
        public string ListUrl { get; set; }
        public bool Enabled { get; set; }
        public string InfoPathTemplate { get; set; }
        public int ItemCount { get; set; }
        public DateTime LastItemUserModifiedDate { get; set; }

        public string strDirName { get; set; }

        public string strLeafName { get; set; }

        public string SiteID { get; set; }

        public string WebID { get; set; }
    }   


    public class Config
    {
        public string Rule { get; set; }
        public Complexity Complexity { get; set; }
        public bool MigrationPathExists { get; set; }
        public bool CustomConnector { get; set; }

    }
    public enum Complexity { Low, Medium, High, Critical };

    public enum Scope { Farm, WebApp, SiteCol };

    public abstract class InfoPathFeature
    {
        #region Useful XNamespace values
        protected static XNamespace xdNamespace = @"http://schemas.microsoft.com/office/infopath/2003";
        protected static XNamespace xsfNamespace = @"http://schemas.microsoft.com/office/infopath/2003/solutionDefinition";
        protected static XNamespace xsf2Namespace = @"http://schemas.microsoft.com/office/infopath/2006/solutionDefinition/extensions";
        protected static XNamespace xsf3Namespace = @"http://schemas.microsoft.com/office/infopath/2009/solutionDefinition/extensions";
        protected static XNamespace xslNamespace = @"http://www.w3.org/1999/XSL/Transform";
        #endregion

        #region Public interface
        public string FeatureName { get { return this.GetType().Name; } }
        // need virtuals for formatting as a string, and an xml, or something ...
        public override string ToString()
        {
            return FeatureName;
        }
        #endregion
        public Complexity Complexity { get; set; }
        public bool CustomConnector { get; set; }
        public bool MigrationPathExists { get; set; }
        public List<Config> Rules { get; set; }

        private void GetConfigFromJSON(JArray configuration)
        {
            foreach (dynamic dataconnection in configuration)
            {
                try
                {
                    Rules.Add(new Config()
                    {
                        Rule = dataconnection.name,
                        Complexity = dataconnection.Complexity,
                        MigrationPathExists = dataconnection.migrationPathExists,
                        CustomConnector = dataconnection.CustomConnector
                    });
                }
                catch(Exception ex)
                {
                    Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                    Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                }
            }
        }

        internal static string LoadDebugJSON()
        {
            string json = string.Empty;
            string dir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            using (StreamReader r = new StreamReader(dir + @"\infopathrules.json"))

            //using (StreamReader r = new StreamReader(@"C:\Users\rmeure\source\repos\InfoPathScrapper\InfoPathScrapper\settings.json"))
            {
                json = r.ReadToEnd();
            }
            return json;
        }
        public InfoPathFeature()
        {
            Rules = new List<Config>();

            var settings = LoadDebugJSON();
            JObject configuration = JObject.Parse(settings);
            dynamic obj = JObject.Parse(settings);

            GetConfigFromJSON((JArray)obj["dataConnections"]["dataConnection"]);
            GetConfigFromJSON((JArray)obj["dataRules"]["dataRule"]);
            GetConfigFromJSON((JArray)obj["dataRules"]["actions"]["action"]);
            GetConfigFromJSON((JArray)obj["controls"]["Control"]);
        }

        #region Abstract method(s)
        /// <summary>
        /// This method returns a commaseparated list of the interesting values a particular feature has collected.
        /// </summary>
        /// <returns></returns>
        public abstract string ToCSV();
        #endregion
    }

    public class Control : InfoPathFeature
    {
        #region Private stuff
        private const string xctName = @"xctname";

        private Control() { }
        #endregion

        #region Public interface
        public string Name { get; private set; }
        public int Count { get; private set; }

        /// <summary>
        /// Instead of logging on feature per control, I do 1 feature per control type along with the number of occurrences
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
        {
            IEnumerable<XElement> allElements = document.Descendants();
            BucketCounter counter = new BucketCounter();
            // collect the control counts
            foreach (XElement element in allElements)
            {
                XAttribute xctAttribute = element.Attribute(xdNamespace + xctName);
                if (xctAttribute != null)
                {
                    counter.IncrementKey(xctAttribute.Value);
                }
            }

            // then create Control objects for each control
            foreach (KeyValuePair<string, int> kvp in counter.Buckets)
            {
                Control c = new Control();
                c.Name = kvp.Key;
                c.Count = kvp.Value;
                yield return c;
            }
            // nothing left
            yield break;
        }

        public override string ToString()
        {
            return FeatureName + ": " + Name + "[" + Count + "]";
            //return "Complexity: " + Complexity + " CustomConnector: " + CustomConnector + " MigrationPathExists: " + MigrationPathExists + " " +  FeatureName + ": " + Name + "[" + Count + "]";
        }

        public override string ToCSV()
        {
            return Name + "," + Count;
        }
        #endregion
    }

    

    public class InfoPathLocation
    {
        public string SiteId { get; set; }
        public string WebId { get; set; }
        public string DocId { get; set; }
        public string DirName { get; set; }
        public string LeafName { get; set; }

    }

    // Data Connection Classes 
    public class DataConnection : InfoPathFeature
    {
        #region Constants
        private const string query = @"query";
        private const string spListConnection = @"sharepointListAdapter";
        private const string spListConnectionRW = @"sharepointListAdapterRW";
        private const string soapConnection = @"webServiceAdapter";
        private const string xmlConnection = @"xmlFileAdapter"; // also used for REST!
        private const string adoConnection = @"adoAdapter";
        #endregion

        #region Public interface
        public string ConnectionType { get; private set; }
        public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
        {

            IEnumerable<XElement> allDataConnections = document.Descendants(xsfNamespace + query);
            foreach (XElement queryElement in allDataConnections)
            {
                yield return ParseDataConnection(queryElement);
            }

            // nothing left
            yield break;
        }

        public string Connection
        { get; private set; }



        public override string ToString()
        {
            return FeatureName + ": " + ConnectionType;
            //return "Complexity: " + Complexity + " CustomConnector: " + CustomConnector + " MigrationPathExists: " + MigrationPathExists + " " + FeatureName + ": " + ConnectionType;
        }

        public override string ToCSV()
        {
            return ConnectionType;
        }
        #endregion

        #region Private helpers
        /// <summary>
        /// This should return DataConnection since every query element represents exactly one connection
        /// In special cases (SPList, Soap) we defer to a subclass to mine more data.
        /// </summary>
        /// <param name="queryElement"></param>
        /// <returns></returns>
        private static DataConnection ParseDataConnection(XElement queryElement)
        {
            XElement dataConnection = queryElement.Element(xsfNamespace + spListConnection);
            if (dataConnection != null)
                return SPListConnection.Parse(dataConnection);
            else if ((dataConnection = queryElement.Element(xsfNamespace + spListConnectionRW)) != null)
                return SPListConnection.Parse(dataConnection);
            else if ((dataConnection = queryElement.Element(xsfNamespace + soapConnection)) != null)
                return SoapConnection.Parse(dataConnection);
            else if ((dataConnection = queryElement.Element(xsfNamespace + xmlConnection)) != null)
                return XmlConnection.Parse(dataConnection);
            else if ((dataConnection = queryElement.Element(xsfNamespace + adoConnection)) != null)
                return AdoConnection.Parse(dataConnection);

            // else just grab the type and log that. Nothing else to do here.
            foreach (XElement x in queryElement.Elements())
            {
                if (dataConnection != null) throw new ArgumentException("More than one adapter found under a query node");
                dataConnection = x;
            }

            if (dataConnection == null) throw new ArgumentException("No adapter found under query node");
            DataConnection dc = new DataConnection();
            dc.ConnectionType = dataConnection.Name.LocalName;
            return dc;
        }
        #endregion
    }

    /// <summary>
    /// Subclass specifically for mining SP List connections
    /// </summary>
    class SPListConnection : DataConnection
    {
        #region Constants
        private const string siteUrlAttribute = @"siteUrl";
        private const string siteURLAttribute = @"siteURL";
        private const string listGuidAttribute = @"sharePointListID";
        private const string sharepointGuidAttribute = @"sharepointGuid";
        private const string submitAllowedAttribute = @"submitAllowed";
        private const string relativeListUrlAttribute = @"relativeListUrl";
        private const string field = @"field";
        private const string typeAttribute = @"type";
        private static string[] validTypes =
        {
            "Counter",
            "Integer",
            "Number",
            "Currency",
            "Text",
            "Choice",
            "Plain",
            "Compatible",
            "FullHTML",
            "DateTime",
            "Boolean",
            "Lookup",
            "LookupMulti",
            "MultiChoice",
            "URL",
            "User",
            "UserMulti",
            "Calculated",
            "Attachments",
            "HybridUser",
        };
        private BucketCounter _bucketCounter;
        #endregion

        private SPListConnection()
        {

            var rule = (from r in Rules
                        where r.Rule == this.GetType().Name
                        select r).First();
            this.Complexity = rule.Complexity;
            this.MigrationPathExists = rule.MigrationPathExists;
            this.CustomConnector = rule.CustomConnector;

            _bucketCounter = new BucketCounter();
            foreach (string type in validTypes)
                _bucketCounter.DefineKey(type);
        }

        #region Public interface

        public string SiteUrl { get; private set; }
        public string ListGuid { get; private set; }
        public string SubmitAllowed { get; private set; }
        public string RelativeListUrl { get; private set; }
        public IEnumerable<KeyValuePair<string, int>> ColumnTypes
        {
            get
            {
                foreach (KeyValuePair<string, int> kvp in _bucketCounter.Buckets)
                    yield return kvp;
                yield break;
            }
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            //sb.Append("Complexity: " + Complexity + " CustomConnector: " + CustomConnector + " MigrationPathExists: " + MigrationPathExists + " ");
            sb.Append(FeatureName).Append(": ");
            sb.Append("{").Append(SiteUrl).Append(", ").Append(RelativeListUrl).Append("} ");
            if (IsV2)
            {
                sb.Append("Types: {");
                foreach (KeyValuePair<string, int> kvp in ColumnTypes)
                {
                    // for humanreadable, just emit nonzero counts
                    if (kvp.Value > 0)
                        sb.Append(kvp.Key).Append("=").Append(kvp.Value).Append(" ");
                }
                sb.Append("}");
            }
            return sb.ToString().Trim();
        }

        public override string ToCSV()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(SiteUrl).Append(",").Append(RelativeListUrl).Append(",");
            if (IsV2)
            {
                foreach (KeyValuePair<string, Int32> kvp in ColumnTypes)
                {
                    sb.Append(kvp.Key).Append(",").Append(kvp.Value).Append(",");
                }
            }
            return sb.ToString();
        }

        private bool IsV2 { get; set; }

        /// <summary>
        /// Selfexplanatory
        /// </summary>
        /// <param name="dataConnection"></param>
        /// <returns></returns>
        public static SPListConnection Parse(XElement dataConnection)
        {
            SPListConnection spl = new SPListConnection();
            spl.IsV2 = dataConnection.Name.LocalName.Equals("sharepointListAdapterRW");
            if (spl.IsV2)
            {
                spl.SiteUrl = dataConnection.Attribute(siteURLAttribute).Value;
                spl.ListGuid = dataConnection.Attribute(listGuidAttribute).Value;
                spl.SubmitAllowed = dataConnection.Attribute(submitAllowedAttribute).Value;
                spl.RelativeListUrl = dataConnection.Attribute(relativeListUrlAttribute).Value;

                // we should also scrape out the queried column types, this shows us what types of data we consume from lists
                foreach (XElement fieldElement in dataConnection.Elements(xsfNamespace + field))
                {
                    string fieldType = fieldElement.Attribute(typeAttribute).Value;
                    spl._bucketCounter.IncrementKey(fieldType);
                }
            }
            else // if (dataConnection.Name.LocalName.Equals("sharepointListAdapter"))
            {
                spl.SiteUrl = dataConnection.Attribute(siteUrlAttribute).Value;
                spl.ListGuid = dataConnection.Attribute(sharepointGuidAttribute).Value;
                spl.SubmitAllowed = dataConnection.Attribute(submitAllowedAttribute).Value;
            }


            return spl;
        }
        #endregion
    }

    /// <summary>
    /// Subclass of DataConnection specifically for Soap web service calls.
    /// </summary>
    class SoapConnection : DataConnection
    {
        #region Constants
        private const string wsdlUrlAttribute = @"wsdlUrl";
        private const string serviceUrlAttribute = @"serviceUrl";
        private const string nameAttribute = @"name";
        private const string operation = @"operation";
        private const string input = @"input";
        private const string sourceAttribute = @"source";

        private const string webServiceAdapterExtension = @"webServiceAdapterExtension";
        private const string refAttribute = @"ref";
        private const string connectoid = @"connectoid";
        private const string udcxExt = @".udcx";
        #endregion

        private SoapConnection()
        {
            var rule = (from r in Rules
                        where r.Rule == this.GetType().Name
                        select r).First();
            this.Complexity = rule.Complexity;
            this.MigrationPathExists = rule.MigrationPathExists;
            this.CustomConnector = rule.CustomConnector;
        }

        #region Public interface
        public string ServiceUrl { get; private set; }
        public string ServiceMethod { get; private set; }

        public static DataConnection Parse(XElement dataConnection)
        {
            XElement udcExtension = null;
            if (IsConnectionUDCX(dataConnection, out udcExtension) && udcExtension != null)
            {
                return UdcConnection.Parse(dataConnection, udcExtension);
            }
            return ParseInternal(dataConnection);
        }

        public override string ToString()
        {
            return FeatureName + ": " + ServiceUrl + "::" + ServiceMethod;
            //return "Complexity: " + Complexity + " CustomConnector: " + CustomConnector + " MigrationPathExists: " + MigrationPathExists + " " + FeatureName + ": " + ServiceUrl + "::" + ServiceMethod;
        }

        public override string ToCSV()
        {
            return ServiceUrl + "," + ServiceMethod;
        }
        #endregion

        #region Private helpers
        private static SoapConnection ParseInternal(XElement dataConnection)
        {
            SoapConnection sc = new SoapConnection();
            XElement op = dataConnection.Element(xsfNamespace + operation);
            if (op == null)
            {
                sc.ServiceUrl = dataConnection.Attribute(wsdlUrlAttribute).Value;
                sc.ServiceMethod = dataConnection.Attribute(nameAttribute).Value;
            }
            else
            {
                sc.ServiceUrl = op.Attribute(serviceUrlAttribute).Value;
                sc.ServiceMethod = op.Attribute(nameAttribute).Value;
                XElement inp = op.Element(xsfNamespace + input);
                if (inp != null && sc.ServiceUrl.Equals(""))
                {
                    sc.ServiceUrl = inp.Attribute(sourceAttribute).Value;
                    sc.ServiceMethod = "?";
                }
            }


            if (sc.ServiceUrl.Equals("")) Console.WriteLine(dataConnection.ToString());
            return sc;
        }


        /// <summary>
        /// Need to find the xsf2:webServiceAdapterExtension node elsewhere in the XDocument
        /// Need to find the one that has ref the same as our connection name 
        /// </summary>
        /// <param name="dataConnection"></param>
        /// <returns></returns>
        private static bool IsConnectionUDCX(XElement dataConnection, out XElement udcExtension)
        {
            udcExtension = null;
            XAttribute name = dataConnection.Attribute(nameAttribute);
            if (name == null) return false;

            string connectionName = name.Value;
            foreach (XElement webServiceExt in dataConnection.Document.Descendants(xsf2Namespace + webServiceAdapterExtension))
            {
                XAttribute refAtt = webServiceExt.Attribute(refAttribute);
                if (refAtt == null) continue; // No name = no match
                if (!refAtt.Value.Equals(connectionName)) continue; // These are not the extensions you are looking for ... *waves hand*

                XElement connect = webServiceExt.Element(xsf2Namespace + connectoid);
                if (connect == null) return false;
                XAttribute source = connect.Attribute(sourceAttribute);
                if (source == null) return false;

                if (Path.GetExtension(source.Value) != udcxExt) return false;

                udcExtension = webServiceExt;
                return true;
            }
            return false;
        }
        #endregion
    }

    /// <summary>
    /// Subclass of DataConnection specifically for UDCX Soap web service calls.
    /// </summary>
    class UdcConnection : DataConnection
    {
        #region Constants
        private const string connectoid = @"connectoid";
        private const string nameAttribute = @"name";
        private const string sourceAttribute = @"source";
        #endregion

        //private UdcConnection() { }

        #region Public interface
        public string SourceUrl { get; private set; }
        public string MethodName { get; private set; }

        public static UdcConnection Parse(XElement dataConnection, XElement udcExtension)
        {
            UdcConnection uc = new UdcConnection();

            XElement connect = udcExtension.Element(xsf2Namespace + connectoid);

            uc.MethodName = connect.Attribute(nameAttribute).Value;
            uc.SourceUrl = connect.Attribute(sourceAttribute).Value;


            return uc;
        }

        protected UdcConnection()
        {
            var rule = (from r in Rules
                        where r.Rule == this.GetType().Name
                        select r).First();
            this.Complexity = rule.Complexity;
            this.MigrationPathExists = rule.MigrationPathExists;
            this.CustomConnector = rule.CustomConnector;
        }

        public override string ToString()
        {
            return FeatureName + ": " + SourceUrl + "::" + MethodName;
            //return "Complexity: " + Complexity + " CustomConnector: " + CustomConnector + " MigrationPathExists: " + MigrationPathExists + " " +  FeatureName + ": " + SourceUrl + "::" + MethodName;
        }

        public override string ToCSV()
        {
            return SourceUrl + "," + MethodName;
        }
        #endregion
    }

    /// <summary>
    /// Identifies a connection to an Xml file
    /// </summary>
    class XmlConnection : DataConnection
    {
        #region Constants
        private const string fileUrlAttribute = @"fileUrl";
        private const string nameAttribute = @"name";
        private const string refAttribute = @"ref";
        private const string xmlFileAdapterExtension = @"xmlFileAdapterExtension";
        private const string isRestAttribute = @"isRest";
        #endregion

        protected XmlConnection()
        {
            var rule = (from r in Rules
                        where r.Rule == this.GetType().Name
                        select r).First();
            this.Complexity = rule.Complexity;
            this.MigrationPathExists = rule.MigrationPathExists;
            this.CustomConnector = rule.CustomConnector;
        }

        #region Public interface
        public string Url { get; private set; }

        public static DataConnection Parse(XElement dataConnection)
        {
            XmlConnection xc = null;
            bool isRest = IsConnectionRest(dataConnection);
            xc = isRest ? new RESTConnection() : new XmlConnection();

            string fileUrl = dataConnection.Attribute(fileUrlAttribute).Value;

            if (!String.IsNullOrEmpty(fileUrl))
            {
                // We have an embedded XmlConnection
                xc.Url = dataConnection.Attribute(fileUrlAttribute).Value;
                return xc;
            }
            else
            {
                // The XmlConnection is stored in a UDCX connection file
                XElement udcExtension = FindChild(dataConnection);
                return UdcConnection.Parse(dataConnection, udcExtension);
            }
        }

        public override string ToString()
        {
            return FeatureName + ": " + Url;
            //return "Complexity: " + Complexity + " CustomConnector: " + CustomConnector + " MigrationPathExists: " + MigrationPathExists + " " + FeatureName + ": " + Url;
        }

        public override string ToCSV()
        {
            return Url;
        }
        #endregion

        #region Private helpers
        private static XElement FindChild(XElement dataConnection)
        {
            XAttribute name = dataConnection.Attribute(nameAttribute);
            if (name == null) return null;

            string connectionName = name.Value;
            foreach (XElement xmlExt in dataConnection.Document.Descendants(xsf2Namespace + xmlFileAdapterExtension))
            {
                XAttribute refAtt = xmlExt.Attribute(refAttribute);
                if (refAtt == null) continue; // No name = no match
                if (!refAtt.Value.Equals(connectionName)) continue; // These are not the extensions you are looking for ... *waves hand*
                return xmlExt;
            }
            return null;
        }

        /// <summary>
        /// Need to find the xsf2:xmlFileAdapterExtension node elsewhere in the XDocument
        /// Need to find the one that has ref the same as our connection name, then look for isRest="[bool]" in there
        /// </summary>
        /// <param name="dataConnection"></param>
        /// <returns></returns>
        private static bool IsConnectionRest(XElement dataConnection)
        {
            XElement xmlExt = FindChild(dataConnection);
            if (xmlExt == null) return false;

            XAttribute isRest = xmlExt.Attribute(isRestAttribute);
            if (isRest == null) return false;

            return isRest.Value.Equals("yes");
        }
        #endregion
    }

    class RESTConnection : XmlConnection
    {
        //private RestConnection()
        //{
        ///	this.Complexity =Complexity.High;
        //	this.CustomConnector = true;
        //}
    }

    class AdoConnection : DataConnection
    {
        #region Constants
        private const string connectionStringAttribute = @"connectionString";
        #endregion

        private AdoConnection()
        {
            var rule = (from r in Rules
                        where r.Rule == this.GetType().Name
                        select r).First();
            this.Complexity = rule.Complexity;
            this.MigrationPathExists = rule.MigrationPathExists;
            this.CustomConnector = rule.CustomConnector;
        }

        #region Public interface
        public string ConnectionString { get; private set; }

        public static AdoConnection Parse(XElement dataConnection)
        {
            AdoConnection ac = new AdoConnection();
            ac.ConnectionString = dataConnection.Attribute(connectionStringAttribute).Value;
            ac.Complexity = Complexity.High;
            return ac;
        }

        public override string ToString()
        {
            return FeatureName + ": " + ConnectionString;
            //  return "Complexity: " + Complexity + " CustomConnector: " + CustomConnector + " MigrationPathExists: " + MigrationPathExists + " " + FeatureName + ": " + ConnectionString;
        }

        public override string ToCSV()
        {
            return ConnectionString;
        }
        #endregion
    }


    // Data Rules Classess

    public class DataRule : InfoPathFeature
    {
        #region Constants
        private const string rule = @"rule";
        #endregion

        #region Public interface
        public string ActionType { get; private set; }

        public DataRule()
        {
            var rule = (from r in Rules
                        where r.Rule == this.GetType().Name
                        select r).First();

            this.Complexity = rule.Complexity;
            this.MigrationPathExists = rule.MigrationPathExists;
            this.CustomConnector = rule.CustomConnector;
        }

        public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
        {
            IEnumerable<XElement> allRules = document.Descendants(xsfNamespace + rule);
            foreach (XElement ruleElement in allRules)
            {
                foreach (DataRule feature in ParseRuleElement(ruleElement))
                    yield return feature;
            }
            yield break;
        }

        public override string ToString()
        {
            return FeatureName + ": " + ActionType;
            //  return "Complexity: " + Complexity + " CustomConnector: " + CustomConnector + " MigrationPathExists: " + MigrationPathExists + " " + FeatureName + ": " + ActionType;
        }

        public override string ToCSV()
        {
            return ActionType;
        }
        #endregion

        #region Private helpers
        private static IEnumerable<DataRule> ParseRuleElement(XElement ruleElement)
        {
            foreach (XElement ruleAction in ruleElement.Elements())
            {
                DataRule feature = new DataRule();
                feature.ActionType = ruleAction.Name.LocalName;
                // we can be any one of many types of rules: dialogbox, assignment, query, submit, switch view 
                // we could parse further if that turns out to be interesting
                yield return feature;
            }
        }
        #endregion
    }

    // Data Validation Class

    public class DataValidation : InfoPathFeature
    {
        #region Constants
        private const string customValidation = @"customValidation";
        #endregion

        #region Public interface
        public string ValidationType { get; private set; }
        private DataValidation()
        {
            var rule = (from r in Rules
                        where r.Rule == this.GetType().Name
                        select r).First();
            this.Complexity = rule.Complexity;
            this.MigrationPathExists = rule.MigrationPathExists;
            this.CustomConnector = rule.CustomConnector;
        }

        public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
        {
            // we don't care about the condition details, just "any custom validation" vs "native cbb"
            IEnumerable<XElement> allValidations = document.Descendants(xsfNamespace + customValidation);
            foreach (XElement validationElement in allValidations)
            {
                DataValidation validation = new DataValidation();
                validation.ValidationType = "Custom validation";
                yield return validation;
            }

            allValidations = document.Descendants(xsf3Namespace + customValidation);
            foreach (XElement validationElement in allValidations)
            {
                DataValidation validation = new DataValidation();
                validation.ValidationType = "Cannot be blank";
                yield return validation;
            }

            yield break;
        }

        public override string ToString()
        {
            return FeatureName + ": " + ValidationType;
            //return "Complexity: " + Complexity + " CustomConnector: " + CustomConnector + " MigrationPathExists: " + MigrationPathExists + " " + FeatureName + ": " + ValidationType;
        }

        public override string ToCSV()
        {
            return ValidationType;
        }
        #endregion
    }

    // Formatting Class

    public class FormattingRule : InfoPathFeature
    {
        #region Constants
        private const string xslAttribute = @"attribute";
        private const string nameAttribute = @"name";
        private const string contentEditable = @"contentEditable"; // readonly
        private const string style = @"style"; // adjusting the look and feel (colors, hiding, etc ...)
        private const string when = @"when";
        #endregion

        private FormattingRule()
        {
            this.Complexity = Complexity.Medium;
            this.CustomConnector = false;
            this.MigrationPathExists = true;
        }

        #region Public interface
        public string FormatType { get; private set; }
        public string SubDetails { get; private set; }

        public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
        {
            IEnumerable<XElement> allElements = document.Descendants(xslNamespace + xslAttribute);
            foreach (XElement element in allElements)
            {
                XAttribute name = element.Attribute(nameAttribute);
                // these are the html attributes that we try to set. 
                // specifically for conditional hide we have to look under a style for an xsl:when with .Text() contains "DISPLAY: none"
                if (name.Value.Equals(contentEditable))
                {
                    FormattingRule rule = new FormattingRule();
                    rule.FormatType = "Readonly";
                    yield return rule;
                }
                else if (name.Value.Equals(style))
                {
                    BucketCounter counter = new BucketCounter();
                    FormattingRule rule = new FormattingRule();
                    rule.FormatType = "Style";
                    // now let's count all the things we're affecting. 
                    // Overloading BucketCounter to filter the noise of multiple touches to same style
                    foreach (XElement xslWhen in element.Descendants(xslNamespace + when))
                    {
                        string[] styles = xslWhen.Value.Split(new char[] { ';' });
                        foreach (string s in styles)
                        {
                            if (s.Trim().StartsWith("caption:")) continue;
                            string affectedStyle = s.Split(':')[0].Trim().ToUpper();
                            counter.IncrementKey(affectedStyle);
                        }
                    }
                    StringBuilder sb = new StringBuilder();
                    foreach (KeyValuePair<string, int> kvp in counter.Buckets)
                    {
                        sb.Append(kvp.Key).Append(" ");
                    }
                    rule.SubDetails = sb.ToString().Trim();
                    yield return rule;
                }
            }

            // nothing left
            yield break;
        }

        public override string ToString()
        {
            return FeatureName + ": " + FormatType + (SubDetails == null ? "" : "," + SubDetails);
            // return "Complexity: " + Complexity + " CustomConnector: " + CustomConnector + " MigrationPathExists: " + MigrationPathExists + " " + FeatureName + ": " + FormatType + (SubDetails == null ? "" : " " + SubDetails);
        }

        public override string ToCSV()
        {
            return FormatType + (SubDetails == null ? "" : "," + SubDetails);
        }
        #endregion
    }

    // Managed Code Class

    class ManagedCode : InfoPathFeature
    {
        private const string enabledAttribute = @"enabled";
        private const string languageAttribute = @"language";
        private const string versionAttribute = @"version";
        private const string managedCode = @"managedCode";

        private ManagedCode()
        {
            this.Complexity = Complexity.High;
            this.CustomConnector = false;
            this.MigrationPathExists = false;
        }

        public string Language { get; private set; }
        public string Version { get; private set; }

        public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
        {
            IEnumerable<XElement> allElements = document.Descendants(xsf2Namespace + managedCode);
            foreach (XElement element in allElements)
            {
                //The enabled attribute is typically null, so check for that. 
                //If the attribute exists, then we can also check the value to make sure there's not an enabled="no" scenario.                
                if (element.Attribute(enabledAttribute) == null
                    || !string.Equals(element.Attribute(enabledAttribute).Value, "yes", StringComparison.InvariantCultureIgnoreCase))
                {
                    continue;
                }

                //Return an object if enabled="yes"
                ManagedCode mc = new ManagedCode();
                mc.Language = element.Attribute(languageAttribute).Value;
                mc.Version = element.Attribute(versionAttribute).Value;
                yield return mc;
            }

            // nothing left
            yield break;
        }

        public override string ToString()
        {
            return FeatureName + ": " + Language + " " + Version;
            //return "Complexity: " + Complexity + " CustomConnector: " + CustomConnector + " MigrationPathExists: " + MigrationPathExists + " " + FeatureName + ": " + Language + " " + Version;
        }

        public override string ToCSV()
        {
            return Language + "," + Version;
        }
    }

    //Mode Class

    class Mode : InfoPathFeature
    {
        private const string modeAttribute = @"mode";
        private const string solutionFormatVersionAttribute = @"solutionFormatVersion";
        private const string solutionMode = @"solutionMode";
        private const string solutionDefinition = @"solutionDefinition";
        private const string solutionPropertiesExtension = @"solutionPropertiesExtension";
        private const string branchAttribute = @"branch";
        private const string runtimeCompatibilityAttribute = @"runtimeCompatibility";
        private const string xDocumentClass = @"xDocumentClass";

        public string ModeName { get; private set; }
        public string Compatibility { get; private set; }

        private Mode() { }

        public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
        {

            Mode m = new Mode();
            string mode = null;

            // look for fancy new modes first (these were new for xsf3 / IP2010)
            IEnumerable<XElement> allModeElements = document.Descendants(xsf3Namespace + solutionMode);
            foreach (XElement element in allModeElements)
            {
                if (mode != null) throw new ArgumentException("Found more than one mode!");
                XAttribute name = element.Attribute(modeAttribute);
                mode = name.Value;
            }

            // and if we didn't find the above, fall back to client v server in xsf2:solutionDefinition
            if (mode == null)
            {
                IEnumerable<XElement> allSolutionDefs = document.Descendants(xsf2Namespace + solutionDefinition);
                foreach (XElement solutionDef in allSolutionDefs)
                {
                    if (mode != null) throw new ArgumentException("Found more than one xsf2:solutionDefition!");
                    XElement extension = solutionDef.Element(xsf2Namespace + solutionPropertiesExtension);
                    if (extension != null && extension.Attribute(branchAttribute) != null && extension.Attribute(branchAttribute).Equals("contentType"))
                    {
                        mode = "Document Information Panel";
                    }
                    else
                    {
                        XAttribute compat = solutionDef.Attribute(runtimeCompatibilityAttribute);
                        mode = compat.Value;
                    }
                }
            }

            // and if we still found nothing, it's a 2003 form and must be client:
            if (mode == null)
                mode = "client";

            m.ModeName = mode;

            string compatibility = null;
            foreach (XElement xDoc in document.Descendants(xsfNamespace + xDocumentClass))
            {
                if (compatibility != null) throw new ArgumentException("Multiple xDocumentClass nodes found!");
                compatibility = xDoc.Attribute(solutionFormatVersionAttribute).Value;
            }
            m.Compatibility = compatibility;

            yield return m;
            yield break;
        }


        public override string ToString()
        {
            return FeatureName + ": " + ModeName + " " + Compatibility;
        }

        public override string ToCSV()
        {
            return ModeName + "," + Compatibility;
        }
    }

    //Product Version Class

    class ProductVersion : InfoPathFeature
    {
        private const string productVersionAttribute = @"productVersion";

        public string Version { get; private set; }

        public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
        {
            //ProductVersion is off the root node, so grab that value and return the value.
            XAttribute attribute = document.Root.Attribute(productVersionAttribute);

            if (attribute != null)
            {
                ProductVersion pv = new ProductVersion();
                pv.Version = attribute.Value;
                yield return pv;
            }

            // nothing left
            yield break;
        }

        public override string ToString()
        {
            return FeatureName + ": " + Version;
        }

        public override string ToCSV()
        {
            return Version;
        }
    }

    //Publish Class

    class PublishUrl : InfoPathFeature
    {
        #region Private stuff
        private const string baseUrl = @"baseUrl";
        private const string relativeUrlBaseAttribute = @"relativeUrlBase";
        private const string publishUrlAttribute = @"publishUrl";
        private const string xDocumentClass = @"xDocumentClass";

        private PublishUrl() { }
        #endregion

        #region Public interface
        public string Publish { get; private set; }
        public string RelativeBase { get; private set; }

        /// <summary>
        /// Instead of logging on feature per control, I do 1 feature per control type along with the number of occurrences
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
        {
            PublishUrl pubRule = new PublishUrl();
            IEnumerable<XElement> allElements = document.Descendants(xsf3Namespace + baseUrl);
            // collect the control counts
            foreach (XElement element in allElements)
            {
                if (pubRule.RelativeBase != null) throw new ArgumentException("Should only see one xsf3:baseUrl node");
                XAttribute pathAttribute = element.Attribute(relativeUrlBaseAttribute);
                if (pathAttribute != null) // this attribute is technically optional per xsf3 spec
                {
                    pubRule.RelativeBase = pathAttribute.Value;
                }
            }

            allElements = document.Descendants(xsfNamespace + xDocumentClass);
            foreach (XElement element in allElements)
            {
                if (pubRule.Publish != null) throw new ArgumentException("Should only see one xsf:xDocumentClass node");
                XAttribute pubUrl = element.Attribute(publishUrlAttribute);
                if (pubUrl != null)
                {
                    pubRule.Publish = pubUrl.Value;
                }
            }

            yield return pubRule;
            // nothing left
            yield break;
        }

        public override string ToString()
        {
            return FeatureName + ": RelativeBase=" + RelativeBase + ", PublishUrl=" + Publish;
        }

        public override string ToCSV()
        {
            return RelativeBase + "," + Publish;
        }
        #endregion
    }

    // InfoPath File Class

    public abstract class InfoPathFile
    {
        private XDocument _xDocument;
        private List<InfoPathFeature> _features;

        #region Public interface
        public CabFileInfo CabFileInfo { get; protected set; }
        public XDocument XDocument { get { InitializeXDocument(); return _xDocument; } }

        /// <summary>
        /// Enumerates the features found in this InfoPathFile.
        /// </summary>
        public IEnumerable<InfoPathFeature> Features
        {
            get
            {
                InitializeFeatures();
                foreach (InfoPathFeature feature in _features)
                    yield return feature;
                yield break;
            }
        }
        #endregion

        #region Abstract methods
        protected abstract IEnumerable<Func<XDocument, IEnumerable<InfoPathFeature>>> FeatureDiscoverers { get; }
        #endregion

        #region Private helpers
        /// <summary>
        /// Loads the XDocument that is the file from the cab. All parseable InfoPath files are xml documents.
        /// </summary>
        private void InitializeXDocument()
        {
            try
            {
                if (_xDocument != null) return;

                Stream content = CabFileInfo.OpenRead();
                // use the 3.5compatible Load API that takes an XmlReader, that way this will work when targeting either .NET3.5 or later
                System.Xml.XmlReader reader = System.Xml.XmlReader.Create(content);
                _xDocument = XDocument.Load(reader);
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }
        }

        /// <summary>
        /// Run all the FeatureDiscoverers for this file type. Each deriving InfoPathFile type
        /// defines what features it *might* contain.
        /// </summary>
        private void InitializeFeatures()
        {
            try
            {
                if (_features != null) return;
                _features = new List<InfoPathFeature>();
                foreach (Func<XDocument, IEnumerable<InfoPathFeature>> discoverer in FeatureDiscoverers)
                    foreach (InfoPathFeature feature in discoverer.Invoke(XDocument))
                        _features.Add(feature);
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }
        }
        #endregion
    }

    // InfoPath Manifestv Class

    public class InfoPathManifest : InfoPathFile
    {
        #region Private members
        private static XNamespace xsfNamespace = @"http://schemas.microsoft.com/office/infopath/2003/solutionDefinition";
        private static XNamespace xsf2Namespace = @"http://schemas.microsoft.com/office/infopath/2006/solutionDefinition/extensions";
        private static XNamespace xsf3Namespace = @"http://schemas.microsoft.com/office/infopath/2009/solutionDefinition/extensions";
        private const string viewNode = @"mainpane";
        private const string viewNameAttribute = @"transform";
        private const string xDocumentClass = @"xDocumentClass";
        private const string uniqueUrn = @"name";

        private XElement _xDocumentNode;
        private List<string> _viewNames;
        private InfoPathManifest() { }
        #endregion

        #region Public stuff
        public List<string> ViewNames { get { InitializeViewNames(); return _viewNames; } }
        public string Name
        {
            get
            {
                InitializeXDocumentNode();

                //InfoPath 2003 forms don't have a Name attribute, so return an empty string instead of throwing an exception.
                return (_xDocumentNode.Attribute(uniqueUrn) == null) ? string.Empty : _xDocumentNode.Attribute(uniqueUrn).Value;
            }
        }
        public static InfoPathManifest Create(CabFileInfo cabFileInfo)
        {
            InfoPathManifest manifest = new InfoPathManifest();
            manifest.CabFileInfo = cabFileInfo;
            return manifest;
        }
        #endregion

        #region Override implementations
        protected override IEnumerable<Func<XDocument, IEnumerable<InfoPathFeature>>> FeatureDiscoverers
        {
            get
            {
                yield return Mode.ParseFeature;
                yield return PublishUrl.ParseFeature;
                yield return DataConnection.ParseFeature;
                yield return ManagedCode.ParseFeature;
                yield return ProductVersion.ParseFeature;
                yield return DataRule.ParseFeature;
                yield return DataValidation.ParseFeature;
                yield break;
            }
        }

        #endregion

        #region Private helpers
        /// <summary>
        /// This is the root node of the manifest.xsf file from which get a few interesting properties.
        /// </summary>
        private void InitializeXDocumentNode()
        {
            if (_xDocumentNode != null) return;
            IEnumerable<XElement> elements = XDocument.Descendants(xsfNamespace + xDocumentClass);
            foreach (XElement element in elements)
            {
                if (_xDocumentNode != null) throw new ArgumentException("Manifest has multiple xDocumentClass nodes");
                _xDocumentNode = element;
            }
        }

        /// <summary>
        /// An xsn can have many resource files in it. The only reliable way to know which ones are views is to 
        /// parse the manifest for those that are called out as such. We just need string names of them because 
        /// the Microsoft.Deployment.Compression.Cab code will happily find them by name.
        /// </summary>
        private void InitializeViewNames()
        {
            if (_viewNames != null) return;
            _viewNames = new List<string>();

            foreach (XElement mainpane in XDocument.Descendants(xsfNamespace + viewNode))
            {
                _viewNames.Add(mainpane.Attribute(viewNameAttribute).Value);
            }
        }
        #endregion

    }

    //InfoPath Template Class


    public class InfoPathTemplate
    {
        #region Members and Basics
        public CabInfo CabInfo { get; private set; }
        public List<InfoPathView> _infoPathViews;
        private InfoPathManifest _infoPathManifest;

        public static InfoPathTemplate CreateTemplate(string path)
        {
            // Lazy Init
            CabInfo cabInfo = new CabInfo(path);
            InfoPathTemplate template = new InfoPathTemplate();
            template.CabInfo = cabInfo;
            return template;
        }
        #endregion


        #region Various Properties we should support
        public List<InfoPathView> InfoPathViews { get { InitializeViews(); return _infoPathViews; } }
        public InfoPathManifest InfoPathManifest { get { InitializeManifest(); return _infoPathManifest; } }
        public IEnumerable<InfoPathFile> FeaturedFiles { get { yield return InfoPathManifest; foreach (InfoPathView view in InfoPathViews) yield return view; yield break; } }
        public IEnumerable<InfoPathFeature> Features { get { foreach (InfoPathFile file in FeaturedFiles) { foreach (InfoPathFeature feature in file.Features) yield return feature; } yield break; } }
        #endregion

        #region Private helpers to compute things
        private void InitializeManifest()
        {
            if (_infoPathManifest != null) return;

            // get the files named manifest.xsf (there should be one)
            IList<CabFileInfo> cbInfos = null;
            try
            {
                cbInfos = CabInfo.GetFiles("manifest.xsf");
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message.ToString());
            }

            // TODO check for corrupt infopath forms
            if (cbInfos == null)
                throw new ArgumentException("Invalid InfoPath xsn");
            if (cbInfos.Count != 1) throw new ArgumentException("Invalid InfoPath xsn");
            _infoPathManifest = InfoPathManifest.Create(cbInfos[0]);
        }

        private void InitializeViews()
        {
            if (_infoPathViews != null) return;
            _infoPathViews = new List<InfoPathView>();

            foreach (string name in InfoPathManifest.ViewNames)
            {
                IList<CabFileInfo> cbInfos = CabInfo.GetFiles(name);
                if (cbInfos.Count != 1) throw new ArgumentException(String.Format("Malformed template file: view {0} not found", name));
                InfoPathView viewFile = InfoPathView.Create(cbInfos[0]);
                _infoPathViews.Add(viewFile);
            }
        }
        #endregion
    }

    //InfoPath View Class

    /// <summary>
    /// Selfexplanatory implementation
    /// </summary>
    public class InfoPathView : InfoPathFile
    {
        #region Private stuff
        private InfoPathView() { }
        #endregion

        #region Public stuff
        public static InfoPathView Create(CabFileInfo cabFileInfo)
        {
            InfoPathView view = new InfoPathView();
            view.CabFileInfo = cabFileInfo;
            return view;
        }
        #endregion

        #region Override implementations
        protected override IEnumerable<Func<XDocument, IEnumerable<InfoPathFeature>>> FeatureDiscoverers
        {
            get
            {
                yield return Control.ParseFeature;
                yield return FormattingRule.ParseFeature;
                yield break;
            }
        }
        #endregion

    }

    public class BucketCounter
    {
        private Dictionary<string, Int32> _dictionary;

        public BucketCounter()
        {
            _dictionary = new Dictionary<string, int>();
        }

        /// <summary>
        /// Creates a key and initializes it to a count of zero.
        /// </summary>
        /// <param name="key"></param>
        public void DefineKey(string key)
        {
            if (!_dictionary.ContainsKey(key))
                _dictionary.Add(key, 0);
        }

        /// <summary>
        /// Make sure we have a counter for the key and increment it. 
        /// Note that this is why I use Int32 objects instead of int
        /// </summary>
        /// <param name="key"></param>
        public void IncrementKey(string key)
        {
            DefineKey(key);
            _dictionary[key]++;
        }

        /// <summary>
        /// For each KeyValuePair, return it. I create new ones here so that 
        /// a caller can't accidentally mess up the values in the Dictionary
        /// The buckets are returned in sortedbykey order because that's more useful than a random order
        // </summary>
        public IEnumerable<KeyValuePair<string, int>> Buckets
        {
            get
            {
                List<string> orderedKeys = new List<string>();
                foreach (string key in _dictionary.Keys)
                    orderedKeys.Add(key);
                orderedKeys.Sort();

                foreach (string key in orderedKeys)
                    yield return new KeyValuePair<string, int>(key, _dictionary[key]);
                yield break;
            }
        }
    }
    */
}

