using System;

namespace BatchAndReport.Models
{
    public class E_ConReport_SMCModels
    {
        public int SMC_R310_60_ID { get; set; }
        public string? Contract_Number { get; set; }
        public DateTime? ContractSignDate { get; set; }
        public string? Contract_Sign_Address { get; set; }
        public string? Contract_Organization { get; set; }
        public string? SignatoryName { get; set; }
        public string? SignatoryPosition { get; set; }
        public string? ContractorType { get; set; }
        public string? ContractorName { get; set; }
        public string? ContractorCompany { get; set; }
        public string? ContractorAddressNo { get; set; }
        public string? ContractorStreet { get; set; }
        public string? ContractorSubDistrict { get; set; }
        public string? ContractorDistrict { get; set; }
        public string? ContractorProvince { get; set; }
        public string? ContractorZipcode { get; set; }
        public string? ContractorSignatoryName { get; set; }
        public string? ContractorSignatoryPosition { get; set; }
        public string? ContractorAuthorize { get; set; }
        public string? CompSetLocation { get; set; }
        public string? RentalBrandName { get; set; }
        public DateTime? ServiceStartDate { get; set; }
        public DateTime? ServiceEndDate { get; set; }
        public int? ServiceTotalYears { get; set; }
        public int? ServiceTotalMonths { get; set; }
        public int? ServiceTotalDays { get; set; }
        public decimal? ServiceFee { get; set; }
        public decimal? VatAmount { get; set; }
        public int? PaymentInstallment { get; set; }
        public int? MaximumDownTimeHours { get; set; }
        public decimal? MaximumDownPercents { get; set; }
        public decimal? PenaltyPerHours { get; set; }
        public int? ServiceFixPerMonths { get; set; }
        public int? ServiceFixStartIn { get; set; }
        public string? ServiceFixStartUnit { get; set; }
        public int? ServiceTimeIn { get; set; }
        public string? ServiceTimeUnit { get; set; }
        public decimal? ServicePenaltyPercent { get; set; }
        public decimal? ContPenaltyPercent { get; set; }
        public decimal? ContPenaltyPerDays { get; set; }
        public string? GuaranteeType { get; set; }
        public decimal? GuaranteeAmount { get; set; }
        public decimal? GuaranteePercent { get; set; }
        public int? NewGuaranteeDays { get; set; }
        public decimal? SubcontractPenaltyPercent { get; set; }
        public int? EnforcementOfFineDays { get; set; }
        public int? OutstandingPeriodDays { get; set; }
        public string? OSMEP_Signer { get; set; }
        public string? OSMEP_Witness { get; set; }
        public string? Contract_Signer { get; set; }
        public string? Contract_Witness { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string? CreateBy { get; set; }
        public DateTime? UpdateDate { get; set; }
        public string? UpdateBy { get; set; }
        public string? Flag_Delete { get; set; }
        public string? LegalEntityRegisNumber { get; set; }
        public DateTime? BusinessRegistrationCertDate { get; set; }
        public bool? AttorneyFlag { get; set; }
        public DateTime? AttorneyLetterDate { get; set; }
        public string? AttorneyLetterNumber { get; set; }
        public string? CitizenId { get; set; }
        public DateTime? CitizenCardRegisDate { get; set; }
        public DateTime? CitizenCardExpireDate { get; set; }
        public int? Request_ID { get; set; }
        public string? Contract_Status { get; set; }
    }

    public class E_ConReport_SMCInstallmentModels
    {
        public int SMC_Inst_ID { get; set; }
        public int SMC_R310_60_ID { get; set; }
        public int PayRound { get; set; }
        public decimal? TotalAmount { get; set; }
        public int RepairMonth { get; set; }
        public string? Flag_Delete { get; set; }
    }
}