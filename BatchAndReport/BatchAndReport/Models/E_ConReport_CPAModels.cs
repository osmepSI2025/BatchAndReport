namespace BatchAndReport.Models
{
    public class E_ConReport_CPAModels
    {
        public int CPA_ID { get; set; }
        public string? Contract_Sign_Address { get; set; }
        public string? Contract_Organization { get; set; }
        public string? SignatoryName { get; set; }
        public string? SignatoryPosition { get; set; }
        public string? ContractorType { get; set; }
        public string? ContractorName { get; set; }
        public string? ContractorAddressNo { get; set; }
        public string? ContractorStreet { get; set; }
        public string? ContractorSubDistrict { get; set; }
        public string? ContractorDistrict { get; set; }
        public string? ContractorProvince { get; set; }
        public string? ContractorZipcode { get; set; }
        public string? ContractorSignatoryName { get; set; }
        public string? ContractorSignatoryPosition { get; set; }
        public DateTime? ContractSignDate { get; set; }
        public string? ContractorAuthorize { get; set; }
        public string? Computer_Model { get; set; }
        public decimal? TotalAmount { get; set; }
        public decimal? VatAmount { get; set; }
        public string? DeliveryLocation { get; set; }
        public int? DeliveryDateIn { get; set; }
        public int? NotiDaysBeforeDelivery { get; set; }
        public int? LocationPrepareDays { get; set; }
        public string? PaymentMethod { get; set; }
        public decimal? AdvancePayment { get; set; }
        public int? PaymentDueDays { get; set; }
        public string? PaymentGuaranteeType { get; set; }
        public decimal? RemainingPaymentAmount { get; set; }
        public string? SaleBankName { get; set; }
        public string? SaleBankBranch { get; set; }
        public string? SaleBankAccountName { get; set; }
        public string? SaleBankAccountNumber { get; set; }
        public string? WarrantyPeriodYears { get; set; }
        public string? WarrantyPeriodMonths { get; set; }
        public int? DaysToRepairAfterNoti { get; set; }
        public int? MaximumDownTimeHours { get; set; }
        public decimal? MaximumDownTimePercent { get; set; }
        public decimal? PenaltyPerHourPercent { get; set; }
        public decimal? PenaltyPerHour { get; set; }
        public int? PenaltyDueDaysIn { get; set; }
        public string? PerformanceGuarantee { get; set; }
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
        public string? OSMEP_Signer { get; set; }
        public string? OSMEP_Witness { get; set; }
        public string? Contract_Signer { get; set; }
        public string? Contract_Witness { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string? CreateBy { get; set; }
        public DateTime? UpdateDate { get; set; }
        public string? UpdateBy { get; set; }
        public string? Flag_Delete { get; set; }
        public string? Contract_Number { get; set; }
        public string? LegalEntityRegisNumber { get; set; }
        public string? CompanyOrganizer { get; set; }
        public bool? AttorneyFlag { get; set; }
        public DateTime? AttorneyLetterDate { get; set; }
        public string? AttorneyLetterNumber { get; set; }
        public string? CitizenId { get; set; }
        public DateTime? CitizenCardRegisDate { get; set; }
        public DateTime? CitizenCardExpireDate { get; set; }
        public DateTime? BusinessRegistrationCertDate { get; set; }
        public string? DeliveryNotifyLocation { get; set; }
        public decimal? PaymentSumAMT { get; set; }
        public int? Request_ID { get; set; }
        public string? Contract_Status { get; set; }
        public string? PaymentGuaranteeTypeOther { get; set; }
        public string? CPAContractNumber { get; set; }
        public List<E_ConReport_SignatoryModels> Signatories { get; set; } = new();

    }

}
