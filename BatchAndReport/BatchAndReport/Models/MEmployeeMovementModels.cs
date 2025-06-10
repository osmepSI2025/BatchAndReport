using BatchAndReport.Models;
using System.Text.Json.Serialization;

namespace BatchAndReport.Models
{
    public class MEmployeeMovementModels
    {
        public int Id { get; set; }

        [JsonPropertyName("employeeId")]
        public string? EmployeeId { get; set; }

        [JsonPropertyName("effectiveDate")]
        public DateTime? EffectiveDate { get; set; }

        [JsonPropertyName("movementTypeId")]
        public string? MovementTypeId { get; set; }

        [JsonPropertyName("movementReasonId")]
        public string? MovementReasonId { get; set; }

        [JsonPropertyName("employeeCode")]
        public string? EmployeeCode { get; set; }

        [JsonPropertyName("employment")]
        public string? Employment { get; set; }

        [JsonPropertyName("employeeStatus")]
        public string? EmployeeStatus { get; set; }

        [JsonPropertyName("employeeTypeId")]
        public string? EmployeeTypeId { get; set; }

        [JsonPropertyName("payrollGroupId")]
        public string? PayrollGroupId { get; set; }

        [JsonPropertyName("companyId")]
        public string? CompanyId { get; set; }

        [JsonPropertyName("businessUnitId")]
        public string? BusinessUnitId { get; set; }

        [JsonPropertyName("positionId")]
        public string? PositionId { get; set; }

        [JsonPropertyName("workLocationId")]
        public string? WorkLocationId { get; set; }

        [JsonPropertyName("calendarGroupId")]
        public string? CalendarGroupId { get; set; }

        [JsonPropertyName("jobTitleId")]
        public string? JobTitleId { get; set; }

        [JsonPropertyName("jobLevelId")]
        public string? JobLevelId { get; set; }

        [JsonPropertyName("jobGradeId")]
        public string? JobGradeId { get; set; }

        [JsonPropertyName("contractStartDate")]
        public DateTime? ContractStartDate { get; set; }

        [JsonPropertyName("contractEndDate")]
        public DateTime? ContractEndDate { get; set; }

        [JsonPropertyName("renewContractCount")]
        public int? RenewContractCount { get; set; }

        [JsonPropertyName("probationDate")]
        public DateTime? ProbationDate { get; set; }

        [JsonPropertyName("probationDuration")]
        public int? ProbationDuration { get; set; }

        [JsonPropertyName("probationResult")]
        public string? ProbationResult { get; set; }

        [JsonPropertyName("probationExtend")]
        public string? ProbationExtend { get; set; }

        [JsonPropertyName("employmentDate")]
        public DateTime? EmploymentDate { get; set; }

        [JsonPropertyName("joinDate")]
        public DateTime? JoinDate { get; set; }

        [JsonPropertyName("occupationDate")]
        public DateTime? OccupationDate { get; set; }

        [JsonPropertyName("terminationDate")]
        public DateTime? TerminationDate { get; set; }

        [JsonPropertyName("terminationReason")]
        public string? TerminationReason { get; set; }

        [JsonPropertyName("terminationSSO")]
        public string? TerminationSSO { get; set; }

        [JsonPropertyName("isBlacklist")]
        public string? IsBlacklist { get; set; }

        [JsonPropertyName("paymentDate")]
        public DateTime? PaymentDate { get; set; }

        [JsonPropertyName("remark")]
        public string? Remark { get; set; }

        [JsonPropertyName("serviceYearAdjust")]
        public int? ServiceYearAdjust { get; set; }

        [JsonPropertyName("supervisorCode")]
        public string? SupervisorCode { get; set; }

        [JsonPropertyName("standardWorkHoursId")]
        public string? StandardWorkHoursID { get; set; }

        [JsonPropertyName("workOperation")]
        public string? WorkOperation { get; set; }
    }
    public class EmployeeMovementResult
    {
        [JsonPropertyName("id")]
        public int? Id { get; set; }

        [JsonPropertyName("employeeId")]
        public string? EmployeeId { get; set; }

        [JsonPropertyName("effectiveDate")]
        public DateTime? EffectiveDate { get; set; }

        [JsonPropertyName("movementTypeId")]
        public string? MovementTypeId { get; set; }

        [JsonPropertyName("movementReasonId")]
        public string? MovementReasonId { get; set; }

        [JsonPropertyName("employeeCode")]
        public string? EmployeeCode { get; set; }

        [JsonPropertyName("employment")]
        public string? Employment { get; set; }

        [JsonPropertyName("employeeStatus")]
        public string? EmployeeStatus { get; set; }

        [JsonPropertyName("employeeTypeId")]
        public string? EmployeeTypeId { get; set; }

        [JsonPropertyName("payrollGroupId")]
        public string? PayrollGroupId { get; set; }

        [JsonPropertyName("companyId")]
        public string? CompanyId { get; set; }

        [JsonPropertyName("businessUnitId")]
        public string? BusinessUnitId { get; set; }

        [JsonPropertyName("positionId")]
        public string? PositionId { get; set; }

        [JsonPropertyName("workLocationId")]
        public string? WorkLocationId { get; set; }

        [JsonPropertyName("calendarGroupId")]
        public string? CalendarGroupId { get; set; }

        [JsonPropertyName("jobTitleId")]
        public string? JobTitleId { get; set; }

        [JsonPropertyName("jobLevelId")]
        public string? JobLevelId { get; set; }

        [JsonPropertyName("jobGradeId")]
        public string? JobGradeId { get; set; }

        [JsonPropertyName("contractStartDate")]
        public DateTime? ContractStartDate { get; set; }

        [JsonPropertyName("contractEndDate")]
        public DateTime? ContractEndDate { get; set; }

        [JsonPropertyName("renewContractCount")]
        public int? RenewContractCount { get; set; }

        [JsonPropertyName("probationDate")]
        public DateTime? ProbationDate { get; set; }

        [JsonPropertyName("probationDuration")]
        public int? ProbationDuration { get; set; }

        [JsonPropertyName("probationResult")]
        public string? ProbationResult { get; set; }

        [JsonPropertyName("probationExtend")]
        public string? ProbationExtend { get; set; }

        [JsonPropertyName("employmentDate")]
        public DateTime? EmploymentDate { get; set; }

        [JsonPropertyName("joinDate")]
        public DateTime? JoinDate { get; set; }

        [JsonPropertyName("occupationDate")]
        public DateTime? OccupationDate { get; set; }

        [JsonPropertyName("terminationDate")]
        public DateTime? TerminationDate { get; set; }

        [JsonPropertyName("terminationReason")]
        public string? TerminationReason { get; set; }

        [JsonPropertyName("terminationSSO")]
        public string? TerminationSSO { get; set; }

        [JsonPropertyName("isBlacklist")]
        public string? IsBlacklist { get; set; }

        [JsonPropertyName("paymentDate")]
        public DateTime? PaymentDate { get; set; }

        [JsonPropertyName("remark")]
        public string? Remark { get; set; }

        [JsonPropertyName("serviceYearAdjust")]
        public int? ServiceYearAdjust { get; set; }

        [JsonPropertyName("supervisorCode")]
        public string? SupervisorCode { get; set; }

        [JsonPropertyName("standardWorkHoursId")]
        public string? StandardWorkHoursID { get; set; }

        [JsonPropertyName("workOperation")]
        public string? WorkOperation { get; set; }
    }


    public class ApiEmployeeMovementResponse
    {
        [JsonPropertyName("results")]
        public EmployeeMovementResult Results { get; set; }
    }
    public class ApiListEmployeeMovementResponse
    {
        [JsonPropertyName("pagination")]
        public Pagination? Pagination { get; set; }
        [JsonPropertyName("results")]
        public List<EmployeeMovementResult?> Results { get; set; } = new();
    }

    public class searchEmployeeMovementModels
    {        
        public string employeeId { get; set; }
    }

   

}
