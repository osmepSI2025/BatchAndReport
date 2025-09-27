using System;
using System.Text.Json.Serialization;

namespace BatchAndReport.Models
{
    public class MEmployeeContractModels
    {
        public int Id { get; set; }
        public bool ContractFlag { get; set; }
        public string EmployeeId { get; set; } = null!;
        public string? EmployeeCode { get; set; }
        public string? NameTh { get; set; }
        public string? NameEn { get; set; }
        public string? FirstNameTh { get; set; }
        public string? FirstNameEn { get; set; }
        public string? LastNameTh { get; set; }
        public string? LastNameEn { get; set; }
        public string? Email { get; set; }
        public string? Mobile { get; set; }
        public DateTime? EmploymentDate { get; set; }
        public DateTime? TerminationDate { get; set; }
        public string? EmployeeType { get; set; }
        public string? EmployeeStatus { get; set; }
        public string? SupervisorId { get; set; }
        public string? CompanyId { get; set; }
        public string? BusinessUnitId { get; set; }
        public string? PositionId { get; set; }
        public string? Salary { get; set; }
        public string? IdCard { get; set; }
        public string? PassportNo { get; set; }
        public string? Address { get; set; }
    }

    public class EmployeeContractResult
    {

        [JsonPropertyName("contractFlag")]
        public bool? ContractFlag { get; set; }
        [JsonPropertyName("employeeId")]
        public string EmployeeId { get; set; } = null!;
        [JsonPropertyName("employeeCode")]
        public string? EmployeeCode { get; set; }
        [JsonPropertyName("nameTh")]
        public string? NameTh { get; set; }
        [JsonPropertyName("nameEn")]
        public string? NameEn { get; set; }
        [JsonPropertyName("firstNameTh")]
        public string? FirstNameTh { get; set; }
        [JsonPropertyName("firstNameEn")]
        public string? FirstNameEn { get; set; }
        [JsonPropertyName("lastNameTh")]
        public string? LastNameTh { get; set; }
        [JsonPropertyName("lastNameEn")]
        public string? LastNameEn { get; set; }
        [JsonPropertyName("email")]
        public string? Email { get; set; }
        [JsonPropertyName("mobile")]
        public string? Mobile { get; set; }
        [JsonPropertyName("employmentDate")]
        public DateTime? EmploymentDate { get; set; }
        [JsonPropertyName("terminationDate")]
        public DateTime? TerminationDate { get; set; }
        [JsonPropertyName("employeeType")]
        public string? EmployeeType { get; set; }
        [JsonPropertyName("employeeStatus")]
        public string? EmployeeStatus { get; set; }
        [JsonPropertyName("supervisorId")]
        public string? SupervisorId { get; set; }
        [JsonPropertyName("companyId")]
        public string? CompanyId { get; set; }
        [JsonPropertyName("businessUnitId")]
        public string? BusinessUnitId { get; set; }
        [JsonPropertyName("positionId")]
        public string? PositionId { get; set; }
        [JsonPropertyName("salary")]
        public string? Salary { get; set; }
        [JsonPropertyName("idCard")]
        public string? IdCard { get; set; }
        [JsonPropertyName("passportNo")]
        public string? PassportNo { get; set; }
    }

    public class ApiEmployeeContractResponse
    {
        [JsonPropertyName("results")]
        public EmployeeContractResult Results { get; set; }
    }

    public class ApiListEmployeeContractResponse
    {
        [JsonPropertyName("pagination")]
        public Pagination? Pagination { get; set; }
        [JsonPropertyName("results")]
        public List<EmployeeContractResult?> Results { get; set; } = new();
    }

    public class SearchEmployeeContractModels
    {
        public int page { get; set; }
        public int perPage { get; set; }
    }
}