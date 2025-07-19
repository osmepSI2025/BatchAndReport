using System;

namespace BatchAndReport.Entities
{
    public class AnnualProcessReviewHistory
    {
        public int AnnualProcessReviewHistoryId { get; set; }
        public int? AnnualProcessReviewId { get; set; }
        public DateTime? Datetime { get; set; }
        public string? StatusCode { get; set; }
        public string? EmployeeId { get; set; }
        public DateTime? CreatedDateTime { get; set; }
        public DateTime? UpdatedDateTime { get; set; }
        public string? CreatedBy { get; set; }
        public string? UpdatedBy { get; set; }
        public bool? IsDeleted { get; set; }
    }
}