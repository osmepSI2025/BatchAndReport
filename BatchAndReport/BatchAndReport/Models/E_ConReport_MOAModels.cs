namespace BatchAndReport.Models
{
    public class E_ConReport_MOAModels
    {
        public long? MOA_ID { get; set; }
        public string? Contract_Number { get; set; }
        public string? ProjectTitle { get; set; }
        public string? OrgName { get; set; }
        public string? OrgCommonName { get; set; }
        public DateTime? Sign_Date { get; set; }
        public string? Requestor { get; set; }
        public string? RequestorPosition { get; set; }
        public string? Org_Requestor { get; set; }
        public string? Org_RequestorPosition { get; set; }
        public string? Contract_Type { get; set; }
        public string? Contract_Type_Other { get; set; }
        public DateTime? Effective_Date { get; set; }
        public string? Office_Loc { get; set; }
        public DateTime? Start_Date { get; set; }
        public DateTime? End_Date { get; set; }
        public decimal? Contract_Value { get; set; }
        public string? Contract_Category { get; set; }
        public string? Contract_Storage { get; set; }
        public string? OSMEP_Signer { get; set; }
        public string? OSMEP_Witness { get; set; }
        public string? Contract_Signer { get; set; }
        public string? Contract_Witness { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string? CreateBy { get; set; }
        public DateTime? UpdateDate { get; set; }
        public string? UpdateBy { get; set; }
        public bool? Flag_Delete { get; set; }
        public string? Request_ID { get; set; }
        public string? Contract_Status { get; set; }
        public bool? NeedAttachCuS { get; set; }
        public bool? AttorneyFlag { get; set; }
        public string? AttorneyLetterNumber { get; set; }
        public string? Organization_Logo { get; set; }
    }
    public class E_ConReport_MOAPoposeModels
    {
        public int? MOAP_ID { get; set; }
        public int? MOA_ID { get; set; }
        public string? Detail { get; set; }

    }
}
