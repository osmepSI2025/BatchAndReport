namespace BatchAndReport.Models
{
    public class E_ConReport_CLAModels
    {
        public int CLA_R309_60_ID { get; set; }
        public string? CLAContractNumber { get; set; }
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
        public string? RentalSysName { get; set; }
        public string? RentalBrandName { get; set; }
        public int? RentalPeriodYear { get; set; }
        public int? RentalPeriodMonth { get; set; }
        public string? SaleBankName { get; set; }
        public string? SaleBankBranch { get; set; }
        public string? SaleBankAccountName { get; set; }
        public string? SaleBankAccountNumber { get; set; }
        public string? DeliveryLocation { get; set; }
        public int? DeliveryDateIn { get; set; }
        public string? NotiLocation { get; set; }
        public int? NotiDaysBeforeDelivery { get; set; }
        public int? LocationDesignDays { get; set; }
        public int? MaintenancePermonth { get; set; }
        public int? MaintenanceInterval { get; set; }
        public int? MaximumDownTimeHours { get; set; }
        public decimal? MaximumDownPercents { get; set; }
        public decimal? PenaltyPerHours { get; set; }
      //  public decimal? NormalTimeFixDays { get; set; }
      //  public decimal? OffTimeFixDays { get; set; }
        public decimal? FixPenaltyPerHours { get; set; }
        public int? FixReplaceCompDays { get; set; }
      //  public decimal? FixReplacePenaltyPerHours { get; set; }
        public int? TrainingPeriodDays { get; set; }
        public int? ComputerManualsCount { get; set; }
        public string? GuaranteeType { get; set; }
        public decimal? GuaranteeAmount { get; set; }
        public decimal? GuaranteePercent { get; set; }
        public int? NewGuaranteeDays { get; set; }
        public int? RespReplaceDays { get; set; }
        public string? TeminationNewMonths { get; set; }
        public decimal? FinePerDaysPercent { get; set; }
        public int? ComputerSendBackDays { get; set; }
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
        public int? RespReplaceYears { get; set; }
        public int? RespReplaceMonth { get; set; }
        public string? Request_ID { get; set; }
        public string? Contract_Status { get; set; }
        public string? GuaranteeTypeOther { get; set; }
        public bool? NeedAttachCuS { get; set; } // bit
        public decimal? NormalTimeFixHours { get; set; }
        public decimal? OffTimeFixHours { get; set; }
        public decimal? FixReplacePenaltyPerDays { get; set; }

    }
    
}
