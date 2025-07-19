using System;

namespace BatchAndReport.Entities
{
    public class AnnualProcessReviewDetail
    {
        public int AnnualProcessReviewDetailId { get; set; }
        public int? AnnualProcessReviewId { get; set; }
        public int? PrevProcessMasterId { get; set; }
        public string? PrevProcessGroupCode { get; set; }
        public string? PrevProcessCode { get; set; }
        public string? PrevProcessName { get; set; }
        public string? ProcessGroupCode { get; set; }
        public string? ProcessCode { get; set; }
        public string? ProcessName { get; set; }
        public string? IsWiFilePath { get; set; }
        public int? ProcessReviewTypeId { get; set; }
        public string? FileUpload { get; set; }
        public bool? IsWorkflow { get; set; }
        public bool? IsCgdControlProcess { get; set; }
        public bool? IsDeleted { get; set; }
        public bool? IsWi { get; set; }
        public bool? PrevIsWorkflow { get; set; }

        public virtual AnnualProcessReview? AnnualProcessReview { get; set; }
    }

}