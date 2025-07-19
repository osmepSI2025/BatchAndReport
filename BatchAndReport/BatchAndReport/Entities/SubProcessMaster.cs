using System;

namespace BatchAndReport.Entities
{
    public class SubProcessMaster
    {
        public int? SubProcessMasterId { get; set; }
        public int? ProcessMasterId { get; set; }
        public int? ProcessDay { get; set; }

        public string? ProcessGroupCode { get; set; }

        public string? ProcessGroupName { get; set; }

        public string? ProcessCode { get; set; }

        public string? ProcessName { get; set; }

        public bool? IsWorkflow { get; set; }

        public bool? IsDigital { get; set; }

        public bool? IsCreateWorkflow { get; set; }

        public string? ProcessTypeCode { get; set; }

        public string? DiagramAttachFile { get; set; }

        public string? ProcessAttachFile { get; set; }

        public string? ApprovalReviewDetail { get; set; }

        public string? StatusCode { get; set; }

        public string? EvaluationStatus { get; set; }

        public string? EvaluationReviewRemark { get; set; }

        public DateTime? CreatedDateTime { get; set; }

        public DateTime? UpdatedDateTime { get; set; }

        public string? CreatedBy { get; set; }

        public string? UpdatedBy { get; set; }

        public bool? IsDeleted { get; set; }

        public int? FiscalYearId { get; set; }

    }
    public class SubProcessMasterWithBU : SubProcessMaster
    {
        public string? OwnerBusinessUnitName { get; set; }
    }
}
