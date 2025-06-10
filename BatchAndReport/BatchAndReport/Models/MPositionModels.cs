using BatchAndReport.Models;
using System.Text.Json.Serialization;

namespace BatchAndReport.Models
{
    public class MPositionModels
    {
        public int Id { get; set; }

        public string? ProjectCode { get; set; }

        public string? PositionId { get; set; }

        public string? TypeCode { get; set; }

        public string? Module { get; set; }

        public string? NameTh { get; set; }

        public string? NameEn { get; set; }

        public string? DescriptionTh { get; set; }

        public string? DescriptionEn { get; set; }
    }
    public class PositionResult
    {
        [JsonPropertyName("projectCode")]
        public string? ProjectCode { get; set; }

        [JsonPropertyName("code")]
        public string? PositionId { get; set; }

        [JsonPropertyName("typeCode")]
        public string? TypeCode { get; set; }

        [JsonPropertyName("module")]
        public string? Module { get; set; }

        [JsonPropertyName("nameTh")]
        public string? NameTh { get; set; }

        [JsonPropertyName("nameEn")]
        public string? NameEn { get; set; }

        [JsonPropertyName("descriptionTh")]
        public string? DescriptionTh { get; set; }

        [JsonPropertyName("descriptionEn")]
        public string? DescriptionEn { get; set; }
    }


    public class ApiPositionResponse
    {
        [JsonPropertyName("results")]
        public PositionResult Results { get; set; }
    }
    public class ApiListPositionResponse
    {
        [JsonPropertyName("pagination")]
        public Pagination? Pagination { get; set; }
        [JsonPropertyName("results")]
        public List<PositionResult?> Results { get; set; } = new();
    }

    public class searchPositionModels
    {        
        public int page { get; set; }
        public int perPage { get; set; }
    }

   

}
