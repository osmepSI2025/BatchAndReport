using BatchAndReport.Models;
using System.Text.Json.Serialization;

namespace BatchAndReport.Models
{
    public class MBusinessUnitModels
    {
        public int Id { get; set; }

        public string? BusinessUnitId { get; set; }

        public string? BusinessUnitCode { get; set; }

        public int? BusinessUnitLevel { get; set; }

        public string? ParentId { get; set; }

        public string? CompanyId { get; set; }

        public DateTime? EffectiveDate { get; set; }

        public string? NameTh { get; set; }

        public string? NameEn { get; set; }

        public string? AbbreviationTh { get; set; }

        public string? AbbreviationEn { get; set; }

        public string? DescriptionTh { get; set; }

        public string? DescriptionEn { get; set; }

        public DateTime? CreateDate { get; set; }
    }

    public class BusinessUnitResult
    {
        [JsonPropertyName("businessUnitId")]
        public string? BusinessUnitId { get; set; }

        [JsonPropertyName("businessUnitCode")]
        public string? BusinessUnitCode { get; set; }

        [JsonPropertyName("businessUnitLevel")]
        public int? BusinessUnitLevel { get; set; }

        [JsonPropertyName("parentId")]
        public string? ParentId { get; set; }

        [JsonPropertyName("companyId")]
        public string? CompanyId { get; set; }

        [JsonPropertyName("effectiveDate")]
        public DateTime? EffectiveDate { get; set; }

        [JsonPropertyName("nameTh")]
        public string? NameTh { get; set; }

        [JsonPropertyName("nameEn")]
        public string? NameEn { get; set; }

        [JsonPropertyName("abbreviationTh")]
        public string? AbbreviationTh { get; set; }

        [JsonPropertyName("abbreviationEn")]
        public string? AbbreviationEn { get; set; }

        [JsonPropertyName("descriptionTh")]
        public string? DescriptionTh { get; set; }

        [JsonPropertyName("descriptionEn")]
        public string? DescriptionEn { get; set; }

        [JsonPropertyName("createDate")]
        public DateTime? CreateDate { get; set; }
    }

    public class ApiBusinessUnitResponse
    {
        [JsonPropertyName("results")]
        public BusinessUnitResult? Results { get; set; }
    }

    public class ApiListBusinessUnitResponse
    {
        [JsonPropertyName("pagination")]
        public Pagination? Pagination { get; set; }

        [JsonPropertyName("results")]
        public List<BusinessUnitResult?> Results { get; set; } = new();
    }

    public class SearchBusinessUnitModels
    {
        public int page { get; set; }
        public int perPage { get; set; }
    }
}
