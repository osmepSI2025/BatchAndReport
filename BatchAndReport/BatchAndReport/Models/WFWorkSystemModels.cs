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
        public string ProcessCode { get; set; }
        public string ProcessName { get; set; }
        public string PrevProcessCode { get; set; }
        public string Department { get; set; }
        public string Workflow { get; set; }
        public string PrevWorkflow { get; set; }
        public string WI { get; set; }
        public string ReviewType { get; set; }
        public string isDigital { get; set; }
    }

}
