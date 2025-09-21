namespace BatchAndReport.Models
{
    public class SubProcessReviewApprovalModels
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

        public string? EmployeeName { get; set; }

        public string? EmployeePosition { get; set; }

        public string? E_Signature { get; set; }

        public string? ActorDetail { get; set; }
    }
}
