namespace BatchAndReport.Models
{
    public class E_ConReport_JOADetailModels
    {
        public string? JOA_ID { get; set; }
        public string? Contract_Number { get; set; }
        public string? Project_Name { get; set; }
        public string? Organization { get; set; }
        public string? Organization_Logo { get; set; }
        public DateTime? Contract_SignDate { get; set; }
        public string? IssueOwner { get; set; }
        public string? IssueOwnerPosition { get; set; }
        public string? JointOfficer { get; set; }
        public string? JointOfficerPosition { get; set; }
        public string? Contract_Type { get; set; }
        public string? Contract_Type_Other { get; set; }
        public bool? AttorneyFlag { get; set; }
        public string? AttorneyLetterNumber { get; set; }
        public DateTime? Grant_Date { get; set; }
        public string? OfficeLoc { get; set; }
        public DateTime? Contract_Start_Date { get; set; }
        public DateTime? Contract_End_Date { get; set; }
        public decimal? Contract_Value { get; set; }
        public string? Contract_Category { get; set; }
        public string? Contract_Storage { get; set; }
        public string? OSMEP_Signer { get; set; }
        public string? Contract_Signer { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string? CreateBy { get; set; }
        public DateTime? UpdateDate { get; set; }
        public string? UpdateBy { get; set; }
        public bool? Flag_Delete { get; set; }
        public string? Request_ID { get; set; }
        public string? Contract_Status { get; set; }
        public List<E_ConReport_SignatoryModels> Signatories { get; set; } = new();
    }
    public class E_ConReport_JOAPoposeModels
    {
        public int? JOAP_ID { get; set; }
        public int? JOA_ID { get; set; }
        public string? Detail { get; set; }
       
    }
    public class E_ConReport_SignatoryModels
    {
        public string? Signatory_Name { get; set; }
        public string? Position { get; set; }
        public string? BU_UNIT { get; set; }
        public string? DS_FILE { get; set; }
        public string? Company_Seal { get; set; }
        public string? Signatory_Type { get; set; }
    }
}
