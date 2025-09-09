namespace BatchAndReport.Models
{
    public class E_ConReport_AMJOAModels
    {
        public string? AMJOA_ID { get; set; }
        public string? Contract_Number { get; set; }
        public string? RefContract_Number { get; set; }
        public string? Contract_Name { get; set; }
        public DateTime? ContractSignDate { get; set; }
        public string? Start_Unit { get; set; }
        public string? Contract_Partner { get; set; }
        public string? Contract_Description { get; set; }
        public string? OSMEP_Signer { get; set; }
        public string? OSMEP_Witness { get; set; }
        public string? Contract_Signer { get; set; }
        public string? Contract_Witness { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string? CreateBy { get; set; }
        public DateTime? UpdateDate { get; set; }
        public string? UpdateBy { get; set; }
        public string? Flag_Delete { get; set; }
        public string? Request_ID { get; set; }
        public string? Contract_Status { get; set; }
        public bool? NeedAttachCuS { get; set; }
        public string? Organization_Logo { get; set; }
        public string? OSMEP_NAME { get; set; }
        public string? OSMEP_POSITION { get; set; }
        public string? CP_S_NAME { get; set; }
        public string? CP_S_POSITION { get; set; }
        public bool? CP_S_AttorneyFlag { get; set; }
        public DateTime? CP_S_AttorneyLetterDate { get; set; }
        public string? CP_S_AttorneyLetterNumbers { get; set; }
    }
    public class E_ConReport_AMJOAObjectModels
    {
        public int? PDPA_ID { get; set; }
        public int? PD_Objectives_ID { get; set; }
        public string? Objective_Description { get; set; }
    }
    public class E_ConReport_AMJOAAgreementListModels
    {
        public int? PD_List_ID { get; set; }
        public string? PDPA_ID { get; set; }
        public string? PD_Detail { get; set; }
    }
}
