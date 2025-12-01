// Controllers/FileController.cs
using BatchAndReport.Models;
using Microsoft.AspNetCore.Mvc;

namespace BatchAndReport.Controllers
{
    public class FileController : Controller
    {
        private readonly IWebHostEnvironment _env;
        public FileController(IWebHostEnvironment env) => _env = env;

        [HttpPost]
        [ValidateAntiForgeryToken]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> Upload(FileUploadModel model)
        {
            if (model.PostedFiles == null || !model.PostedFiles.Any())
            {
                TempData["StatusMessage"] = "No files uploaded.";
                TempData["IsSuccess"] = false;
                return RedirectToPage("/MultiUploadFile");
            }

            var targetFolder = string.IsNullOrWhiteSpace(model.ProcessInstanceID)
                ? Path.Combine(_env.WebRootPath, "Document", "ImportContract")
                : Path.Combine(_env.WebRootPath, "Document", "ImportContract", model.ProcessInstanceID.Trim());

            Directory.CreateDirectory(targetFolder);

            var saved = new List<object>();
            var skipped = new List<object>();

            foreach (var file in model.PostedFiles)
            {
                if (file == null || file.Length <= 0)
                {
                    skipped.Add(new { name = file?.FileName, reason = "Empty file" });
                    continue;
                }

                var originalName = Path.GetFileName(file.FileName ?? "");
                var ext = Path.GetExtension(originalName);
                var looksLikePdf = string.Equals(ext, ".pdf", StringComparison.OrdinalIgnoreCase);

                bool headerIsPdf = false;
                try
                {
                    using var head = file.OpenReadStream();
                    if (head.CanRead)
                    {
                        var sig = new byte[4];
                        var read = await head.ReadAsync(sig, 0, sig.Length);
                        headerIsPdf = read == 4 && sig[0] == 0x25 && sig[1] == 0x50 && sig[2] == 0x44 && sig[3] == 0x46; // %PDF
                    }
                }
                catch { }

                if (!(looksLikePdf || headerIsPdf))
                {
                    skipped.Add(new { name = originalName, reason = "Not a PDF" });
                    continue;
                }

                var safeFileName = SanitizeFileName(originalName);
                if (string.IsNullOrWhiteSpace(Path.GetFileNameWithoutExtension(safeFileName)))
                {
                    safeFileName = $"file-{DateTime.UtcNow:yyyyMMddHHmmssfff}.pdf";
                }

                if (!safeFileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                    safeFileName = Path.ChangeExtension(safeFileName, ".pdf");

                safeFileName = MakeUnique(targetFolder, safeFileName);

                var filePath = Path.Combine(targetFolder, safeFileName);

                try
                {
                    using var fs = new FileStream(filePath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
                    await file.CopyToAsync(fs);

                    var baseSegment = string.IsNullOrWhiteSpace(model.ProcessInstanceID)
                        ? "/Document/ImportContract"
                        : $"/Document/ImportContract/{Uri.EscapeDataString(model.ProcessInstanceID.Trim())}";

                    var url = $"{baseSegment}/{Uri.EscapeDataString(safeFileName)}";

                    saved.Add(new
                    {
                        originalName,
                        savedName = safeFileName,
                        size = file.Length,
                        url
                    });
                }
                catch (Exception ex)
                {
                    skipped.Add(new { name = originalName, reason = "Save failed: " + ex.Message });
                }
            }

            TempData["StatusMessage"] = $"Upload completed: {saved.Count} file(s) saved, {skipped.Count} file(s) skipped.";
            TempData["IsSuccess"] = saved.Any();
            return RedirectToPage("/MultiUploadFile");
        }

        public IActionResult Index()
        {
            var xmodel = new FileUploadModel();
            return View(xmodel);
        }

        // --- Helpers ---

        private static string SanitizeFileName(string name)
        {
            // เอาเฉพาะชื่อไฟล์จริง
            var justName = Path.GetFileName(name ?? string.Empty);

            // ตัดอักขระต้องห้ามของไฟล์ระบบ
            var invalid = Path.GetInvalidFileNameChars();
            var cleaned = new string(justName.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());

            // กันความยาวเกินไป
            if (cleaned.Length > 180)
            {
                var ext = Path.GetExtension(cleaned);
                var stem = Path.GetFileNameWithoutExtension(cleaned);
                cleaned = stem.Substring(0, Math.Max(1, 180 - (ext?.Length ?? 0))) + ext;
            }
            return cleaned;
        }

        private static string MakeUnique(string folder, string fileName)
        {
            var name = Path.GetFileNameWithoutExtension(fileName);
            var ext = Path.GetExtension(fileName);
            var candidate = fileName;
            var i = 1;
            while (System.IO.File.Exists(Path.Combine(folder, candidate)))
            {
                candidate = $"{name} ({i++}){ext}";
            }
            return candidate;
        }
    }
}
