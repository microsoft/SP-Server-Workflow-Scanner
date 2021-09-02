using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common
{
    public class WorkflowActionAnalysis
    {

        public WorkflowActionAnalysis()
        {
            this.WorkflowActions = new List<string>();
            this.UnsupportedActions = new List<string>();
        }

        public List<string> WorkflowActions { get; set; }
        public int ActionCount { get; set; }
        public List<string> UnsupportedActions { get; set; }
        public int UnsupportedAccountCount { get; set; }
    }
}
