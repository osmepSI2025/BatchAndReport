using System;

namespace BatchAndReport.Models
{
    public class E_ConReport_PDSAModels
    {
        public int PDSA_ID { get; set; }
        public string? Contract_Number { get; set; }
        public string? Project_Name { get; set; }
        public string? Contract_Organization { get; set; }
        public string? Master_Contract_Number { get; set; }
        public DateTime? Master_Contract_Sign_Date { get; set; }
        public string? ContractPartyName { get; set; }
        public string? ContractPartyCommonName { get; set; }
        public string? ContractPartyType { get; set; }
        public string? ContractPartyType_Other { get; set; }
        public string? Contract_Category { get; set; }
        public string? Contract_Storage { get; set; }
        public int? RetentionPeriodDays { get; set; }
        public int? IncidentNotifyPeriod { get; set; }
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
    }

    public class PDSA_LegalBasisSharing
    {
        public int LglShare_ID { get; set; }
        public int? PDSA_ID { get; set; }
        public string? Detail { get; set; }
        public string? Owner { get; set; }
        public string? Flag_Delete { get; set; }
    }
    public class PDSA_Shared_Data
    {
        public int SharePD_ID { get; set; }
        public int? PDSA_ID { get; set; }
        public string? Detail { get; set; }

        public string? Objective { get; set; }
        public string? Owner { get; set; }
        public string? Flag_Delete { get; set; }
    }

}