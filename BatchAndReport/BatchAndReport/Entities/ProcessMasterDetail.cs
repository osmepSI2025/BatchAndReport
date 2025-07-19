using System;
using System.Collections.Generic;

namespace BatchAndReport.Entities
{
    public class ProcessMasterDetail
    {
        public int ProcessMasterDetailId { get; set; }
        public int? ProcessMasterId { get; set; }
        public string? ProcessTypeCode { get; set; }
        public string? ProcessGroupCode { get; set; }
        public string? ProcessGroupName { get; set; }
        public DateTime? CreatedDateTime { get; set; }
        public DateTime? UpdatedDateTime { get; set; }
        public string? CreatedBy { get; set; }
        public string? UpdatedBy { get; set; }
        public int? FiscalYearId { get; set; }
        public bool? IsDeleted { get; set; }
    }


}
