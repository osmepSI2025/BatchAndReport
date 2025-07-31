namespace BatchAndReport.Models
{
    public class WFCreateProcessStatusModels
    {
        public int? subProcessMasterId { get; set; }
        public string? FiscalYearDesc { get; set; }
        public string? BUNameTh { get; set; }
        public string? ProcessType { get; set; }
        public string? ProcessGroupCode { get; set; }
        public string? ProcessGroupName { get; set; }
        public string? ProcessCode { get; set; }
        public string? ProcessName { get; set; }
        public int? FiscalYearId { get; set; }
        public string? ProcessTypeCode { get; set; }
        public string? Status { get; set; }
        public List<WFCreateProcessStatusModels>? CreateProcessStatusModels { get; set;}
    }
}
