namespace BatchAndReport.Models
{
    public class LookUpModels
    {
        public int Id { get; set; }
        public string? LookupCode { get; set; }
        public string? LookupType { get; set; }
        public string? LookupValue { get; set; }
        public string? FlagDelete { get; set; }
        public DateTime? CreateDate { get; set; }
        public string? CreateBy { get; set; }
        public DateTime? UpdateDate { get; set; }
        public string? UpdateBy { get; set; }
    }
}
