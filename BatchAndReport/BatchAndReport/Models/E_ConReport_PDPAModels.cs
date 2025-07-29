namespace BatchAndReport.Models
{
    public class E_ConReport_PDPAModels
    {
        public int PDPA_ID { get; set; }
        public string Contract_Number { get; set; }
        public string Project_Name { get; set; }
        public string Contract_Organization { get; set; }
        public string Master_Contract_Number { get; set; }
        public DateTime? Master_Contract_Sign_Date { get; set; }
        public string ContractPartyName { get; set; }
        public string ContractPartyCommonName { get; set; }
        public string ContractPartyType { get; set; }
        public string ContractPartyType_Other { get; set; }
        public string OSMEP_ScopeRightsDuties { get; set; }
        public string Contract_Ref_Name { get; set; }
        public DateTime? Start_Date { get; set; }
        public DateTime? End_Date { get; set; }
        public string Contract_Category { get; set; }
        public string Contract_Storage { get; set; }
        public string Objectives { get; set; }
        public string Objectives_Other { get; set; }
        public string RecordFreq { get; set; }
        public string RecordFreqUnit { get; set; }
        public int? RetentionPeriodDays { get; set; }
        public int? IncidentNotifyPeriod { get; set; }
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
    }

    public class E_ConReport_PDPAObjectModels {
        public int? PDPA_ID { get; set; }
        public int? PD_Objectives_ID { get; set; }
        public string? Objective_Description { get; set; }
    }
    public class E_ConReport_PDPAAgreementListModels
    {
        public int? PD_List_ID { get; set; }
        public string? PDPA_ID { get; set; }
        public string? PD_Detail { get; set; }
    }
}
