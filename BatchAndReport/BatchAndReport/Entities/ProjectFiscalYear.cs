using System;
using System.Collections.Generic;

namespace BatchAndReport.Entities
{   
    public class ProjectFiscalYear
    {
        public int FiscalYearId { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public string? FiscalYearDesc { get; set; }
        public DateTime? CreateDate { get; set; }
        public string? CreateBy { get; set; }
        public DateTime? UpdateDate { get; set; }
        public string? UpdateBy { get; set; }
        public bool? ActiveFlag { get; set; }
        public string? StartEndDateDesc { get; set; }
        public virtual ICollection<AnnualProcessReview> AnnualProcessReviews { get; set; }
    }
}
