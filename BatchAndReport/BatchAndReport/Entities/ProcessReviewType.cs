using System;

namespace BatchAndReport.Entities
{
    public class ProcessReviewType
    {
        public int ProcessReviewTypeId { get; set; }
        public string? ProcessReviewTypeName { get; set; }
        public DateTime? CreatedDateTime { get; set; }
        public DateTime? UpdatedDateTime { get; set; }
        public string? CreatedBy { get; set; }
        public string? UpdatedBy { get; set; }
        public bool? IsActive { get; set; }
        public bool? IsDeleted { get; set; }
    }
}