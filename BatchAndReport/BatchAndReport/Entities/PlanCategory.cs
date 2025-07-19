using System;
using System.Collections.Generic;

namespace BatchAndReport.Entities
{
    public class PlanCategory
    {
        public int PlanCategoriesId { get; set; }
        public string? PlanCategoriesName { get; set; }
        public DateTime? CreatedDateTime { get; set; }
        public DateTime? UpdatedDateTime { get; set; }
        public string? CreatedBy { get; set; }
        public string? UpdatedBy { get; set; }
        public bool? IsActive { get; set; }
        public bool? IsDeleted { get; set; }

        public virtual ICollection<PlanCategoriesDetail> PlanCategoriesDetails { get; set; } = new List<PlanCategoriesDetail>();
    }
}
