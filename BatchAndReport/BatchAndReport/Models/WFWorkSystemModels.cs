using BatchAndReport.Models;
using System.Text.Json.Serialization;

namespace BatchAndReport.Models
{
    public class WorkSystemModels
    {
        public string FiscalYear { get; set; }
        public string PreviousYear { get; set; }
        public List<ProcessGroupDto> ProcessGroups { get; set; } = new();
        public List<ProcessDetailDto> ProcessDetails { get; set; } = new();
    }

    public class ProcessGroupDto
    {
        public string PROCESS_GROUP_CODE { get; set; }
        public string PROCESS_GROUP_NAME { get; set; }
    }

    public class ProcessDetailDto
    {
        public string ProcessYear { get; set; }
        public string ProcessCode { get; set; }
        public string ProcessName { get; set; }
        public string PrevProcessCode { get; set; }
        public string Department { get; set; }
        public string Workflow { get; set; }
        public string PrevWorkflow { get; set; }
        public string WI { get; set; }
        public string ReviewType { get; set; }
        public string isDigital { get; set; }

        public string PROCESS_GROUP_CODE { get; set; }
        public string PROCESS_GROUP_NAME { get; set; }
        public string FISCAL_YEAR_DESC { get; set; }

        public int PROCESS_MASTER_DETAIL_ID { get; set; }
        

    }
    public class Wf_tasklistModels
    {
        public int? WFTaskListID { get; set; }
        public int? WF_ID { get; set; }
        public string STATUS { get; set; }
        public int? Request_ID { get; set; }
        public string WF_TYPE { get; set; }
        public DateTime COMPLETEON { get; set; }

        public string OWNER { get; set; }
    }
}
