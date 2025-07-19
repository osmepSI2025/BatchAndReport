using System;
using System.Collections.Generic;

namespace BatchAndReport.Entities
{
    public class PlanCategoriesDetail
    {
        public int PlanCategoriesDetailId { get; set; }
        public int? PlanCategoriesId { get; set; }
        public string? BusinessUnitId { get; set; }
        public DateTime? CreatedDateTime { get; set; }
        public DateTime? UpdatedDateTime { get; set; }
        public string? CreatedBy { get; set; }
        public string? UpdatedBy { get; set; }
        public bool? IsActive { get; set; }
        public bool? IsDeleted { get; set; }
        public string? Objective { get; set; }

        public virtual PlanCategory? PlanCategory { get; set; }
    }
}
