using System;

namespace BatchAndReport.Models
{
    public class E_ConReport_PMLModels
    {
        public int PML_R314_60_ID { get; set; }
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
        public DateTime? ContractorAuthDate { get; set; }
        public string? ContractorAuthNumber { get; set; }
        public string? RentalCopierBrand { get; set; }
        public string? RentalCopierModel { get; set; }
        public string? RentalCopierNumber { get; set; }
        public int? RentalCopierAmount { get; set; }
        public int? RentalYears { get; set; }
        public int? RentalMonths { get; set; }
        public DateTime? RentalStartDate { get; set; }
        public DateTime? RentalEndDate { get; set; }
        public decimal? RatePerUnit { get; set; }
        public decimal? RateTotal { get; set; }
        public int? EstCopiesPerMonth { get; set; }
        public int? IfNotCopiesAmount { get; set; }
        public decimal? CopiesRate { get; set; }
        public string? SaleBankName { get; set; }
        public string? SaleBankBranch { get; set; }
        public string? SaleBankAccountName { get; set; }
        public string? SaleBankAccountNumber { get; set; }
        public string? DeliveryLocation { get; set; }
        public string? DeliveryType { get; set; }
        public int? TotalDay { get; set; }
        public DateTime? DeliveryDate { get; set; }
        public string? NotiSendLocation { get; set; }
        public int? NotiDaysBeforeDelivery { get; set; }
        public int? RespReplaceDays { get; set; }
        public int? MaintenancePermonth { get; set; }
        public int? MaintenanceInterval { get; set; }
        public int? CopierFixDays { get; set; }
        public int? ReplaceFixDays { get; set; }
        public decimal? FinePerDays { get; set; }
        public string? GuaranteeType { get; set; }
        public decimal? GuaranteeAmount { get; set; }
        public decimal? GuaranteePercent { get; set; }
        public int? NewGuaranteeDays { get; set; }
        public int? TeminationReplaceDays { get; set; }
        public decimal? LateFinePerDays { get; set; }
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
        public bool? AttorneyFlag { get; set; }
        public DateTime? AttorneyLetterDate { get; set; }
        public string? AttorneyLetterNumber { get; set; }
        public string? CitizenId { get; set; }
        public DateTime? CitizenCardRegisDate { get; set; }
        public DateTime? CitizenCardExpireDate { get; set; }
        public int? CopierSendBackDays { get; set; }
        public int? Request_ID { get; set; }
        public string? Contract_Status { get; set; }
    }
}
