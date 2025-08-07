using System;

namespace BatchAndReport.Models
{
    public class E_ConReport_SLAModels
    {
        public int? SLA_R308_60_ID { get; set; }
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
        public decimal? TotalAmount { get; set; }
        public decimal? VatAmount { get; set; }
        public string? SWRight_1 { get; set; }
        public int? SWRight_1_Detail { get; set; }
        public string? SWRight_2 { get; set; }
        public int? SWRight_2_Detail_1 { get; set; }
        public int? SWRight_2_Detail_2 { get; set; }
        public string? SWRight_3 { get; set; }
        public int? SWRight_3_Detail { get; set; }
        public string? SWRight_4 { get; set; }
        public string? SWRight_5 { get; set; }
        public string? SWRight_5_Detail { get; set; }
        public DateTime? SWExpiry_Date { get; set; }
        public int? TotalYears { get; set; }
        public int? TotalMonths { get; set; }
        public int? TotalDays { get; set; }
        public string? DeliveryLocation { get; set; }
        public int? DeliveryDateIn { get; set; }
        public string? NotiLocation { get; set; }
        public int? NotiDaysBeforeDelivery { get; set; }
        public string? PaymentMethod { get; set; }
        public int? PaymentInstallment { get; set; }
        public decimal? LastPayRound { get; set; }
        public string? SaleBankName { get; set; }
        public string? SaleBankBranch { get; set; }
        public string? SaleBankAccountName { get; set; }
        public string? SaleBankAccountNumber { get; set; }
        public int? BackupQty { get; set; }
        public string? ComputerModel { get; set; }
        public string? ComputerBrand { get; set; }
        public string? BuyerAddressNo { get; set; }
        public string? BuyerStreet { get; set; }
        public string? BuyerSubDistrict { get; set; }
        public string? BuyerDistrict { get; set; }
        public string? BuyerProvince { get; set; }
        public string? BuyerZipcode { get; set; }
        public string? WarrantyPeriodYears { get; set; }
        public string? WarrantyPeriodMonths { get; set; }
        public int? DaysToRepairIn { get; set; }
        public string? GuaranteeType { get; set; }
        public decimal? GuaranteeAmount { get; set; }
        public decimal? GuaranteePercent { get; set; }
        public int? NewGuaranteeDays { get; set; }
        public int? TrainingPeriodDays { get; set; }
        public int? ComputerManualsCount { get; set; }
        public string? TeminationNewMonths { get; set; }
        public int? ReturnDaysIn { get; set; }
        public decimal? FinePerDaysPercent { get; set; }
        public int? EnforcementOfFineDays { get; set; }
        public int? OutstandingPeriodDays { get; set; }
        public int? ComputerSendBackDays { get; set; }
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
        public string? GuaranteeTypeOther { get; set; }

      
    }

    public class E_ConReport_SLAInstallmentModels
    {
        public int SLA_Inst_ID { get; set; }
        public int SLA_R308_60_ID { get; set; }
        public int PayRound { get; set; }
        public decimal? TotalAmount { get; set; }
        public int UseMonth { get; set; }
        public string? Flag_Delete { get; set; }
    }
}