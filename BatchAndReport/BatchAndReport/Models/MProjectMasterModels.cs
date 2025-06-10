using System.Text.Json.Serialization;

namespace BatchAndReport.Models
{
    public class MProjectMasterModels
    {
        public int? ProjectMasterId { get; set; }
        public int? KeyId { get; set; }
        public string? ProjectName { get; set; }

        public decimal? BudgetAmount { get; set; }

        public string? Issue { get; set; }

        public string? Strategy { get; set; }

        public string? FiscalYear { get; set; }
    }

    public class ProjectMasterResult
    {
        [JsonPropertyName("DATA_P1")]
        public string? DATA_P1 { get; set; }
        [JsonPropertyName("DATA_P2")]
        public string? DATA_P2 { get; set; }
        [JsonPropertyName("DATA_P3")]
        public string? DATA_P3 { get; set; }
        [JsonPropertyName("DATA_P4")]
        public DateTime? DATA_P4 { get; set; }
        [JsonPropertyName("DATA_P5")]
        public DateTime? DATA_P5 { get; set; }
        [JsonPropertyName("DATA_P6")]
        public string? DATA_P6 { get; set; }
        [JsonPropertyName("DATA_P7")]
        public string? DATA_P7 { get; set; }
        [JsonPropertyName("DATA_P8")]
        public string? DATA_P8 { get; set; }
        [JsonPropertyName("DATA_P9")]
        public string? DATA_P9 { get; set; }
        [JsonPropertyName("DATA_P10")]
        public string? DATA_P10 { get; set; }
        [JsonPropertyName("DATA_P11")]
        public string? DATA_P11 { get; set; }
        [JsonPropertyName("DATA_P12")]
        public decimal? DATA_P12 { get; set; }
        [JsonPropertyName("DATA_P13")]
        public decimal? DATA_P13 { get; set; }
    }

    public class ApiProjectMasterResponse
    {
        [JsonPropertyName("results")]
        public ProjectMasterResult? Results { get; set; }
    }

    public class ApiListProjectMasterResponse
    {
        [JsonPropertyName("pagination")]
        public Pagination? Pagination { get; set; }

        [JsonPropertyName("results")]
        public List<ProjectMasterResult?> Results { get; set; } = new();
    }

    public class SearchProjectMasterModels
    {
        public int page { get; set; }
        public int perPage { get; set; }
    }

    public class ApiResponseReturnProjectModels
    {
        public int? StatusCode { get; set; }
        public string? Message { get; set; }
        public Dictionary<string, ProjectMasterResult> Data { get; set; } = new();
    }
}
