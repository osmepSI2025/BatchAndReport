namespace BatchAndReport.Models
{
    public class E_ConReport_RelatedDocumentsModels
    {
        public int Document_ID { get; set; }
        public int Contract_ID { get; set; }
        public string? Contract_Type { get; set; }
        public string? DocumentTitle { get; set; }
        public string? Required_Flag { get; set; }
        public string? FilePath { get; set; }
        public int PageAmount { get; set; }
        public string? Flag_Delete { get; set; }
        public string? File_Name { get; set; }
        public string? File_Location { get; set; }
    }
}
