namespace BatchAndReport.Models
{
    public class WFProcessDetailModels
    {
        public int FiscalYear { get; set; } = 0;

        public List<ProcessGroupItem>? CoreProcesses { get; set; } = new();
        public List<ProcessGroupItem>? SupportProcesses { get; set; } = new();
        public string? FiscalYearPrevious { get; set; }
        public string[]? ReviewDetails { get; set; }
        public List<string>? PrevProcesses { get; set; }
        public List<string>? CurrentProcesses { get; set; }
        public List<string>? ControlActivities { get; set; }
        public List<string>? WorkflowProcesses { get; set; }
        public string[]? ApproveRemarks { get; set; }
        public string? Approver1Name { get; set; }
        public string? Approver2Name { get; set; }
        public string? Approve1Date { get; set; }
        public string? Approve2Date { get; set; }
        public string? Approver1Position { get; set; }
        public string? Approver2Position { get; set; }
    }

    public class ProcessGroupItem
    {
        public string? ProcessGroupCode { get; set; }
        public string? ProcessGroupName { get; set; }
    }

}
