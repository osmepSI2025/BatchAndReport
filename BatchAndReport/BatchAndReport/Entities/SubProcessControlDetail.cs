using System;

namespace BatchAndReport.Entities
{
    public class SubProcessControlDetail
    {
        public int? SubProcessControlDetailId { get; set; }

        public int? SubProcessMasterId { get; set; }

        public string? ProcessControlCode { get; set; }

        public string? ProcessControlActivity { get; set; }

        public string? ProcessControlDetail { get; set; }

        public int? ProcessControlDay { get; set; }

        public DateTime? CreatedDateTime { get; set; }

        public DateTime? UpdatedDateTime { get; set; }

        public string? CreatedBy { get; set; }

        public string? UpdatedBy { get; set; }

        public bool? IsDeleted { get; set; }
    }
}
