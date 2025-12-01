namespace BatchAndReport.Models
{
    public class UploadResultViewModel
    {
        public string Message { get; set; } = "";
        public int Total { get; set; }
        public int SavedCount { get; set; }
        public int SkippedCount { get; set; }
        public List<UploadedFileInfo> SavedFiles { get; set; } = new();
        public List<UploadedFileInfo> SkippedFiles { get; set; } = new();
    }

    public class UploadedFileInfo
    {
        public string Name { get; set; }
        public string Reason { get; set; }
        public string Url { get; set; }
        public long Size { get; set; }
    }
}
