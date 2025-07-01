public class SMESummaryProjectModels
{
    public int? IssueId { get; set; }           //
    public string? IssueName { get; set; }      // ST.TOPIC
    public int? ProjectCount { get; set; }      // COUNT(SPR.PROJECT_CODE)
    public decimal? Budget { get; set; }        // SUM(SPR.BUDGET_AMOUNT)
}
public class SMEStrategyDetailModels
{
    public int? StrategyId { get; set; }
    public string? Topic { get; set; }
    public string? StrategyDesc { get; set; }
    public string? DepartmentCode { get; set; }
    public string? Department { get; set; }
    public string? ProjectName { get; set; }
    public decimal? BudgetAmount { get; set; }
    public string? ProjectStatus { get; set; }
}
