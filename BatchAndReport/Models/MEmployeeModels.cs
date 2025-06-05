using BatchAndReport.Models;
using System.Text.Json.Serialization;

namespace BatchAndReport.Models
{
    public class MEmployeeModels
    {
        public int Id { get; set; }

        public string? EmployeeId { get; set; }

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
    }
    public class EmployeeResult
    {
        [JsonPropertyName("employeeId")]
        public string? EmployeeId { get; set; }

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
    }


    public class ApiEmployeeResponse
    {
        [JsonPropertyName("results")]
        public EmployeeResult Results { get; set; }
    }
    public class ApiListEmployeeResponse
    {
        [JsonPropertyName("pagination")]
        public Pagination? Pagination { get; set; }
        [JsonPropertyName("results")]
        public List<EmployeeResult?> Results { get; set; } = new();
    }

    public class searchEmployeeModels
    {        
        public int page { get; set; }
        public int perPage { get; set; }
    }

   

}
