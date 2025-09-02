using System.ComponentModel.DataAnnotations;

namespace BatchAndReport.Models
{
    public class FileUploadModel
    {
        [Required]
        public List<IFormFile>? PostedFiles { get; set; }
        public string? ProcessInstanceID { get; set; }  // hidden field จากหน้า .cshtml
    }
}
