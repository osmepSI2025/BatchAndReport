using System;

namespace BatchAndReport.Entities
{
    public class Evaluation
    {
        public int? EvaluationId { get; set; }

        public int? SubProcessMasterId { get; set; }

        public string? EvaluationDesc { get; set; }

        public string? PerformanceResult { get; set; }

        public DateTime? CreatedDateTime { get; set; }

        public DateTime? UpdatedDateTime { get; set; }

        public string? CreatedBy { get; set; }

        public string? UpdatedBy { get; set; }

        public bool? IsDeleted { get; set; }
    }
}
