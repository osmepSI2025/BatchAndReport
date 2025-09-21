namespace BatchAndReport.Models
{
    namespace BatchAndReport.Models
    {
        public class E_ConReport_JDCAModels
        {
            public int JDCA_ID { get; set; }
            public string Contract_Number { get; set; }
            public string Project_Name { get; set; }
            public string Contract_Party_Name { get; set; }
            public string Contract_Party_Abb_Name { get; set; }
            public string Contract_Party_Type { get; set; }
            public string Contract_Party_Type_Other { get; set; }
            public string MOU_Name { get; set; }
            public string Master_Contract_Number { get; set; }
            public DateTime? Master_Contract_Sign_Date { get; set; }
            public string Contract_Category { get; set; }
            public string Contract_Storage { get; set; }
            public string OSMEP_ContRep { get; set; }
            public string OSMEP_ContRep_Contact { get; set; }
            public string OSMEP_DPO { get; set; }
            public string OSMEP_DPO_Contact { get; set; }
            public string CP_ContRep { get; set; }
            public string CP_ContRep_Contact { get; set; }
            public string CP_DPO { get; set; }
            public string CP_DPO_Contact { get; set; }
            public string OSMEP_Signer { get; set; }
            public string OSMEP_Witness { get; set; }
            public string Contract_Signer { get; set; }
            public string Contract_Witness { get; set; }
            public DateTime? CreatedDate { get; set; }
            public string CreateBy { get; set; }
            public DateTime? UpdateDate { get; set; }
            public string UpdateBy { get; set; }
            public bool Flag_Delete { get; set; }
            public string Request_ID { get; set; }
            public string Contract_Status { get; set; }

            public string Organization_Logo { get; set; }
            
        }
    }
    public class E_ConReportJDCA_JointPurpModels
    {
        public int? JP_ID { get; set; }
        public int? JDCA_ID { get; set; }
        public string? Detail { get; set; }
    }
    public class E_ConReport_JDCA_PurpMeansModels
    {
        public int? PM_ID { get; set; }
        public int? JDCA_ID { get; set; }
        public string? Detail { get; set; }
    }
    public class E_ConReport_JDCA_SubProcessActivitiesModels
    {
        public int? SubPA_ID { get; set; }
        public int? JDCA_ID { get; set; }
        public string? Activity { get; set; }
        public string? LegalBasis { get; set; }
        public string? PersonalData { get; set; }
        public string? Owner { get; set; }

    }
}
