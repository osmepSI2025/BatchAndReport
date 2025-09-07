namespace BatchAndReport.Models
{
    public class E_ConReport_NDAModels
    {
        public int NDA_ID { get; set; }
        public string Contract_Number { get; set; }
        public string Contract_Party_Name { get; set; }
        public DateTime? Sign_Date { get; set; }
        public string OSMEP_Signatory { get; set; }
        public string OSMEP_Position { get; set; }
        public string CP_Signatory { get; set; }
        public string CP_Position { get; set; }
        public string Contract_Type { get; set; }
        public string Contract_Type_Other { get; set; }
        public string OfficeLoc { get; set; }
        public string Contract_Category { get; set; }
        public string Contract_Storage { get; set; }
        public string Ref_Name { get; set; }
        public string EnforcePeriods { get; set; }
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
        public string? OSMEP_NAME { get; set; }
        public bool? CP_S_AttorneyFlag { get; set; }
        public DateTime? Grant_Date { get; set; }
        public string? AttorneyLetterNumber { get; set; }
        public string? OSMEP_POSITION { get; set; }
        public bool? AttorneyFlag { get; set; }
        public string? CP_S_NAME { get; set; }
        public string? CP_S_POSITION { get; set; }
        public DateTime? CP_S_AttorneyLetterDate { get; set; }
    }

    public class E_ConReport_NDAConfidentialTypeModels
    {
        public int? Conf_ID { get; set; }
        public int? NDA_ID { get; set; }
        public string? Detail { get; set; }
    }
    public class E_ConReport_NDA_RequestPurposeModels
    {
        public int? RP_ID { get; set; }
        public int? NDA_ID { get; set; }
        public string? Detail { get; set; }
    }
}
