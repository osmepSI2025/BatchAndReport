namespace BatchAndReport.Models
{
    public class WorkflowPoint
    {
        public int Id { get; set; }
        public string WorkflowTitle { get; set; }
        public string Department { get; set; }
        public int EditNumber { get; set; }
        public DateTime EditDate { get; set; }
        public string PageNumber { get; set; }
        public string Indicators { get; set; }
        public List<WorkflowApproval> Approvals { get; set; }
        public List<WorkflowHistory> HistoryEdits { get; set; }
    }

    public class WorkflowApproval
    {
        public int Level { get; set; }
        public string SignText { get; set; }
        public string FullName { get; set; }
        public string Position { get; set; }
    }

    public class WorkflowHistory
    {
        public int EditNumber { get; set; }
        public DateTime EditDate { get; set; }
        public string Description { get; set; }
    }
}
