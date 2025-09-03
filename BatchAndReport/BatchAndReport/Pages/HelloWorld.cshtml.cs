using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BatchAndReport.Pages
{
    public class HelloWorldModel : PageModel
    {
        private readonly IWebHostEnvironment _env;

        public HelloWorldModel(IWebHostEnvironment env)
        {
            _env = env;
        }

        [BindProperty]
        public List<IFormFile> PostedFiles { get; set; } = new List<IFormFile>();

        [BindProperty]
        public string? ProcessInstanceID { get; set; } = string.Empty;

        public string? StatusMessage { get; set; }
        public bool IsSuccess { get; set; }
        public string Message { get; private set; } = "Hello World!";

        public void OnGet()
        {
            PostedFiles = new List<IFormFile>();
            ProcessInstanceID = string.Empty;
            StatusMessage = null;
        }

        public async Task<IActionResult> OnPostUploadAsync()
        {
            if (PostedFiles == null || !PostedFiles.Any())
            {
                StatusMessage = "No files uploaded.";
                IsSuccess = false;
                return Page();
            }

            // Use IWebHostEnvironment to get the correct web root path
            var folderPath = Path.Combine(_env.WebRootPath, "Document", "ImportContract");
            var targetFolder = string.IsNullOrWhiteSpace(ProcessInstanceID)
                ? folderPath
                : Path.Combine(folderPath, ProcessInstanceID);

            try
            {
                Directory.CreateDirectory(targetFolder);

                foreach (var file in PostedFiles)
                {
                    if (file.Length <= 0) continue;

                    var safeFileName = Path.GetFileName(file.FileName);
                    var filePath = Path.Combine(targetFolder, safeFileName);

                    using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
                    await file.CopyToAsync(stream);
                }

                StatusMessage = $"Successfully uploaded {PostedFiles.Count} file(s).";
                IsSuccess = true;
            }
            catch (Exception ex)
            {
                StatusMessage = $"An error occurred during upload: {ex.Message}";
                IsSuccess = false;
                // Log the full exception for debugging
                // Example: _logger.LogError(ex, "File upload failed.");
            }

            return Page();
        }
    }
}
