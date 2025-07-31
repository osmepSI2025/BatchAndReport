using System;
using System.Collections.Generic;

namespace BatchAndReport.Entities
{
    public class TempProcessMasterDetail
    {
        public int ProcessMasterDetailId { get; set; }
        public int? ProcessMasterId { get; set; }
        public string? ProcessTypeCode { get; set; }
        public string? ProcessGroupCode { get; set; }
        public string? ProcessGroupName { get; set; }
        public DateTime? CreatedDateTime { get; set; }

        public string? CreatedBy { get; set; }

        public int? FiscalYearId { get; set; }
        public string? USER_PROCESS_REVIEW_NAME { get; set; }
        
    }


}
