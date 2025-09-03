using System.ComponentModel.DataAnnotations;

namespace BatchAndReport.Models
{
    public class ProcessMasterModels
    {
        public int PROCESS_MASTER_ID { get; set; }
        public string? USER_PROCESS_REVIEW_NAME { get; set; }
        public string? VISION_NAME { get; set; }
        public DateTime? CREATED_DATETIME { get; set; }
        public DateTime? UPDATED_DATETIME { get; set; }
        public string? CREATED_BY { get; set; }
        public string? UPDATED_BY { get; set; }
        public int? FISCAL_YEAR_ID { get; set; }
        public bool? IS_DELETED { get; set; }
    }
    
}
