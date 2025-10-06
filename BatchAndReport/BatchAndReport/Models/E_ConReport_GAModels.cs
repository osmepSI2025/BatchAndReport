namespace BatchAndReport.Models
{
    public class E_ConReport_GAModels
    {
        public int GA_ID { get; set; }
        public string Contract_Number { get; set; }
        public DateTime? ContractSignDate { get; set; }
        public string SignAddress { get; set; }
        public string OrganizationName { get; set; }
        public string SignatoryName { get; set; }
        public string SignatoryPosition { get; set; }
        public string TaxID { get; set; }
        public string ContractPartyName { get; set; }
        public string RegType { get; set; }
        public string RegOrganization { get; set; }
        public string HQLocationAddressNo { get; set; }
        public string HQLocationStreet { get; set; }
        public string HQLocationSubDistrict { get; set; }
        public string HQLocationDistrict { get; set; }
        public string HQLocationProvince { get; set; }
        public string HQLocationZipCode { get; set; }
        public string RegEmail { get; set; }
        public string RegPersonalName { get; set; }
        public string RegIdenID { get; set; }
        public decimal? GrantAmount { get; set; }
        public DateTime? GrantStartDate { get; set; }
        public DateTime? GrantEndDate { get; set; }
        public string SpendingPurpose { get; set; }
        public string OSMEP_Signer { get; set; }
        public string OSMEP_Witness { get; set; }
        public string Contract_Signer { get; set; }
        public string Contract_Witness { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string CreateBy { get; set; }
        public DateTime? UpdateDate { get; set; }
        public string UpdateBy { get; set; }
        public string Flag_Delete { get; set; }
        public string? Request_ID { get; set; }
        public string Contract_Status { get; set; }
    }
    
}
