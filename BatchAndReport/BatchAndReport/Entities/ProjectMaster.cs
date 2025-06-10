using System;
using System.Collections.Generic;

namespace BatchAndReport.Entities
{
    public class ProjectMaster
    {
        public int ProjectMasterId { get; set; }      // Not nullable
        public int? KeyId { get; set; }
        public string ProjectName { get; set; }       // Not nullable
        public decimal BudgetAmount { get; set; }     // Not nullable
        public string? Issue { get; set; }
        public string? Strategy { get; set; }
        public string? FiscalYear { get; set; }
    }
}
