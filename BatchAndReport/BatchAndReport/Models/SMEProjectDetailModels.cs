public class SMEProjectDetailModels
{
    public string? ProjectCode { get; set; }
    public string? ProjectName { get; set; }
    public string? MinistryName { get; set; }
    public string? DepartmentCode { get; set; }
    public string? DepartmentName { get; set; }
    public string? ActivityName { get; set; }
    public int FiscalYear { get; set; }
  
    public decimal BudgetAmount { get; set; }
    public decimal? BudgetAmountApprove { get; set; }
    public string? StrategyDesc { get; set; }
    public List<string>? OperationArea { get; set; }
    public int? Score { get; set; }
    public string? ProjectStatus { get; set; }
    public string? ProjectStatusName { get; set; }
    public string? ProjectRationale { get; set; }
    public string? ProjectObjective { get; set; }
    public string? TargetGroup { get; set; }
    public List<string>? TargetGroups { get; set; }
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }

    // 🔽 Additional fields
    public List<string>? Plans { get; set; }
    public string? OwnerName { get; set; }
    public string? OwnerPosition { get; set; }
    public string? OwnerPhone { get; set; }
    public string? OwnerMobile { get; set; }
    public string? OwnerEmail { get; set; }
    public string? OwnerLineId { get; set; }
    public string? ContactName { get; set; }
    public string? ContactPosition { get; set; }
    public string? ContactPhone { get; set; }
    public string? ContactMobile { get; set; }
    public string? ContactEmail { get; set; }
    public string? ContactLineId { get; set; }

    public List<string>? PromotionStrategies { get; set; }
    public string? Activities { get; set; }
    public string? ProjectFocus { get; set; }
    public List<string>? IndustrySector { get; set; }
    public string? Timeline { get; set; }
    public string? OrgPartner { get; set; }
    public string? RoleDescription { get; set; }
    public List<string>? SoftPowers { get; set; }
    public List<Indicator>? OutputIndicators { get; set; }
    public List<Indicator>? OutcomeIndicators { get; set; }
    public List<Strategy>? Strategies { get; set; }
    public string? AdditionalNotes { get; set; }

    // หาปี จาก FiscalYear Master
    public string? FiscalYearDesc { get; set; }
    public string? MinistryId { get; set; }
    public string? IS_BUDGET_USED_FLAG { get; set; }
    public string? SME_ISSUE_CODE { get; set; }

    public string? TARGET_BRANCH_LIST { get; set; }
    public string? DaysDiff { get; set; }
    public string? Partner_Name { get; set; }
    public string? BUDGET_SOURCE_NAME { get; set; }

    public string? Topic { get; set; }
    public string? STRATEGY_DESC { get; set; }
    

}

public class Indicator
{
    public string? Name { get; set; }
    public string? Target { get; set; }
    public string? Unit { get; set; }
    public string? Method { get; set; }
}
public class Strategy
{
    public string? StrategyId { get; set; }
    public string? Topic { get; set; }
    public string? StrategyDesc { get; set; }
}

public class OwnerAndContactDetailsModels
{
    // 🔽 Additional fields
    public List<string>? Plans { get; set; }
    public string? OwnerName { get; set; }
    public string? OwnerPosition { get; set; }
    public string? OwnerPhone { get; set; }
    public string? OwnerMobile { get; set; }
    public string? OwnerEmail { get; set; }
    public string? OwnerLineId { get; set; }
    public string? ContactName { get; set; }
    public string? ContactPosition { get; set; }
    public string? ContactPhone { get; set; }
    public string? ContactMobile { get; set; }
    public string? ContactEmail { get; set; }
    public string? ContactLineId { get; set; }
}