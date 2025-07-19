using System;

namespace BatchAndReport.Entities
{
    public class SubProcessReviewApproval
    {
        public int SubProcessReviewApprovalId { get; set; }

        public int? SubProcessMasterId { get; set; }

        public string? ApprovalTypeCode { get; set; }

        public string? EmployeePositionId { get; set; }

        public string? EmployeeId { get; set; }

        public DateTime? CreatedDateTime { get; set; }

        public DateTime? UpdatedDateTime { get; set; }

        public string? CreatedBy { get; set; }

        public string? UpdatedBy { get; set; }

        public bool? IsDeleted { get; set; }
    }
}
