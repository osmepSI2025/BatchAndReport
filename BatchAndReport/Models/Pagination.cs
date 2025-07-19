using System.Text.Json.Serialization;

namespace BatchAndReport.Models
{
    public class Pagination
    {
        [JsonPropertyName("totalRecords")]
        public int TotalRecords { get; set; }

        [JsonPropertyName("currentPage")]
        public int CurrentPage { get; set; }

        [JsonPropertyName("pageCount")]
        public int PageCount { get; set; }

        [JsonPropertyName("nextPage")]
        public int? NextPage { get; set; }

        [JsonPropertyName("prevPage")]
        public int? PrevPage { get; set; }
    }
}
