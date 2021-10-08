using System;
using System.Collections.Generic;

namespace Common
{
    public class WorkflowScanResult : Scan
    {

        public WorkflowScanResult()
        {
            this.UsedActions = new List<string>();
            this.UnsupportedActionsInFlow = new List<string>();
            this.UsedTriggers = new List<string>();
            this.LastSubscriptionEdit = DateTime.MinValue;
            this.LastDefinitionEdit = DateTime.MinValue;
        }

        public string CreatedBy { get; set; }

        public string ModifiedBy { get; set; }

        public DateTime CreatedOn { get; set; }

        public DateTime ModifiedOn { get; set; }

        public DateTime LastRun { get; set; }

        public string ListUrl { get; set; }
 
        public string ListTitle { get; set; }

        public Guid ListId { get; set; }

        public string ContentTypeName { get; set; }
        public string ContentTypeId { get; set; }

        public Guid DefinitionId { get; set; }

        public Guid SubscriptionId { get; set; }

        /// <summary>
        /// 2010 or 2013 workflow engine
        /// </summary>
        public string Version { get; set; }

        public bool IsOOBWorkflow { get; set; }

        /// <summary>
        /// Site, List, ContentType
        /// </summary>
        public string Scope { get; set; }

        /// <summary>
        /// Site, List or Universal workflow (2013) or N/A (2010)
        /// </summary>
        public string RestrictToType { get; set; }

        public bool Enabled { get; set; }

        /// <summary>
        /// Calculation showing if one should consider upgrading this workflow
        /// </summary>
        public bool ConsiderUpgradingToFlow
        {
            get
            {
                if ((Scope == "List" || Scope == "ContentType" || Scope == "Site") &&
                    Enabled && !IsOOBWorkflow && HasSubscriptions)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public string DefinitionName { get; set; }

        public string DefinitionDescription { get; set; }

        public string SubscriptionName { get; set; }

        public bool HasSubscriptions { get; set; }

        public string AssociationData { get; set; }
        public int ActionCount { get; set; }

        public List<string> UsedActions
        {
            get; set;
        }
        // calculates Flow Mapping Percentage 
        // item.Value.ToFLowMappingPercentage
        public int ToFLowMappingPercentage
        {
            get
            {
                if (ActionCount == 0)
                {
                    return -1;
                }
                else
                {
                    return (int)(((double)(ActionCount - UnsupportedActionCount) / (double)ActionCount) * 100);
                }
            }
        }

        public int UnsupportedActionCount { get; set; }

        public List<string> UnsupportedActionsInFlow { get; set; }

        public List<string> UsedTriggers { get; set; }

        public DateTime LastSubscriptionEdit { get; set; }

        public DateTime LastDefinitionEdit { get; set; }
        public bool AllowManual { get; set; }
        public bool AutoStartChange { get; set; }
        public bool AutoStartCreate { get; set; }
    }
}
