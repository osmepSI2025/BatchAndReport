namespace BatchAndReport.Models
{
    public class AnnualProcessReviewModels
    {
        public int AnnualProcessReviewId { get; set; }
        public string? ProcessReviewDetail { get; set; }
        public string? ProcessBackground { get; set; }
        public string? OwnerBusinessUnitId { get; set; }
        public string? StatusCode { get; set; }
        public string? Detail { get; set; }
        public DateTime? CreatedDateTime { get; set; }
        public DateTime? UpdatedDateTime { get; set; }
        public string? CreatedBy { get; set; }
        public string? UpdatedBy { get; set; }
        public int? FiscalYearId { get; set; }
        public bool? IsDeleted { get; set; }
        public bool? IsDraft { get; set; }
        public string? ApproveRemark { get; set; }
    }
}
