using System;

namespace BatchAndReport.Entities
{
    public class SubProcessMasterHistory
    {
        public int? SubProcessMasterHistoryId { get; set; }

        public int? SubProcessMasterId { get; set; }

        public string? ProcessMasterHistoryType { get; set; }

        public string? EditDetail { get; set; }

        public DateTime? DateTime { get; set; }

        public string? StatusCode { get; set; }

        public string? EmployeeId { get; set; }

        public DateTime? CreatedDateTime { get; set; }

        public DateTime? UpdatedDateTime { get; set; }

        public string? CreatedBy { get; set; }

        public string? UpdatedBy { get; set; }

        public bool? IsDeleted { get; set; }
    }
}
