namespace BatchAndReport.Models
{
    public class relate_LawsModels
    {
        public int RELATED_LAWS_ID { get; set; }                // int, primary key
        public int SUB_PROCESS_MASTER_ID { get; set; }          // int, foreign key
        public string RELATED_LAWS_DESC { get; set; }           // nvarchar(MAX), law description
        public DateTime? CREATED_DATETIME { get; set; }         // datetime, created timestamp
        public DateTime? UPDATED_DATETIME { get; set; }         // datetime, updated timestamp
        public string CREATED_BY { get; set; }                  // nvarchar(50), created by user
        public string UPDATED_BY { get; set; }                  // nvarchar(50), updated by user
        public bool IS_DELETED { get; set; }                    // bit, soft delete flag
    }
}