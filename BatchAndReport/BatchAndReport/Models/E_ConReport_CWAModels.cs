using System;

public class E_ConReport_CWAModels
{
    public int CWA_ID { get; set; }
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
    public string? WorkName { get; set; }
    public string? GuaranteeType { get; set; }
    public decimal? GuaranteeAmount { get; set; }
    public decimal? GuaranteePercent { get; set; }
    public int? GuaranteePaymentPeriod { get; set; }
    public string? PaymentMethod { get; set; }
    public decimal? PrepaidAmount { get; set; }
    public decimal? PrepaidPercents { get; set; }
    public string? PrepaidGuaranteeType { get; set; }
    public decimal? PrepaidDeductPercent { get; set; }
    public DateTime? WorkStartDate { get; set; }
    public DateTime? WorkEndDate { get; set; }
    public int? WarrantyPeriodYears { get; set; }
    public int? WarrantyPeriodMonths { get; set; }
    public decimal? SubcontractPenaltyPercent { get; set; }
    public decimal? FinePerDays { get; set; }
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
    public int? DaysToRepairIn { get; set; }
    public int? Request_ID { get; set; }
    public string? Contract_Status { get; set; }
    public decimal? Install_PayAMT { get; set; }
    public decimal? Install_PayVat { get; set; }
    public int? Install_Num { get; set; }
    public string? BankName { get; set; }
    public string? BankBranch { get; set; }
    public string? BankAccountName { get; set; }
    public string? BankAccountNumber { get; set; }
}

public class CWA_Installment
{
    public int CWA_Inst_ID { get; set; }
    public int CWA_ID { get; set; }
    public int PayRound { get; set; }
    public decimal? TotalAmount { get; set; }
    public string? WorkName { get; set; }
    public DateTime? DeliverDate { get; set; }
    public bool Flag_Inst_Final { get; set; }
}