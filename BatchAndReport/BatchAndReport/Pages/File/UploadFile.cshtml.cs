using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BatchAndReport.Models;

namespace BatchAndReport.Pages.File
{
    public class UploadFileModel : PageModel
    {
      //  private readonly IWebHostEnvironment _env;
        private readonly ILogger<UploadFileModel> _logger;
        public UploadFileModel
            (
            //IWebHostEnvironment env,
            
            ILogger<UploadFileModel> logger)
        {
         //   _env = env;
            _logger = logger;
        }

        [BindProperty]
        public FileUploadModel FileUpload { get; set; } = new();
        public string? StatusMessage { get; set; }
        public bool IsSuccess { get; set; }

        public void OnGet()
        {
            // Set initial values to prevent null reference exceptions on first load.
            FileUpload.ProcessInstanceID = string.Empty;
        }

        public async Task<IActionResult> OnPostAsync()
        {
            try
            {
                if (FileUpload.PostedFiles == null || !FileUpload.PostedFiles.Any())
                {
                    StatusMessage = "No files selected for upload.";
                    IsSuccess = false;
                    return Page();
                }

                // Use Directory.GetCurrentDirectory() to get the base path
                var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "ImportContract");
                var targetFolder = string.IsNullOrWhiteSpace(FileUpload.ProcessInstanceID)
                    ? folderPath
                    : Path.Combine(folderPath, FileUpload.ProcessInstanceID);

                try
                {
                    Directory.CreateDirectory(targetFolder);

                    foreach (var file in FileUpload.PostedFiles)
                    {
                        if (file.Length <= 0) continue;

                        var safeFileName = Path.GetFileName(file.FileName);
                        var filePath = Path.Combine(targetFolder, safeFileName);

                        using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
                        await file.CopyToAsync(stream);
                    }

                    StatusMessage = $"Successfully uploaded {FileUpload.PostedFiles.Count} file(s).";
                    IsSuccess = true;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "An error occurred during upload.");
                    StatusMessage = $"An error occurred during upload: {ex.Message}";
                    IsSuccess = false;
                }

                return Page();

            }
            catch (System.Exception ex)
            {
                _logger.LogError(ex, "An unexpected error occurred in OnPostAsync.");
                var msg = ex.Message;
                StatusMessage = $"An unexpected error occurred: {msg}";
                IsSuccess = false;
                return Page();
            }
        }
    }
}
