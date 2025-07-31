namespace BatchAndReport.Models
{
    public class WFProcessResultByIndicatorModels
    {
        public string? BUNameTh { get; set; }
        public string? ProcessType { get; set; }
        public string? ProcessGroupCode { get; set; }
        public string? ProcessCode { get; set; }
        public string? ProcessName { get; set; }
        public string? EvaluationDesc { get; set; }
        public string? PerformanceResult { get; set; }
        public string? FiscalYearDesc { get; set; }
        public List<WFProcessResultByIndicatorModels>? ProcessResultByIndicators { get; set;}
    }
}
