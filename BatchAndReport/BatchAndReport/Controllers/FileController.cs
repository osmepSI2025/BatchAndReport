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
                return BadRequest(new { message = "No files uploaded." });

            // โฟลเดอร์ปลายทาง
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

                // 1) ตรวจสอบนามสกุล
                var originalName = Path.GetFileName(file.FileName ?? "");
                var ext = Path.GetExtension(originalName);
                var looksLikePdf = string.Equals(ext, ".pdf", StringComparison.OrdinalIgnoreCase);

                // 2) ตรวจสอบ header ว่าเป็น %PDF หรือไม่
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
                catch { /* ignore header read error, will fall back to ext check */ }

                if (!(looksLikePdf || headerIsPdf))
                {
                    skipped.Add(new { name = originalName, reason = "Not a PDF" });
                    continue;
                }

                // 3) sanitize ชื่อไฟล์
                var safeFileName = SanitizeFileName(originalName);
                if (string.IsNullOrWhiteSpace(Path.GetFileNameWithoutExtension(safeFileName)))
                {
                    // ถ้าชื่อว่าง ให้ตั้งใหม่
                    safeFileName = $"file-{DateTime.UtcNow:yyyyMMddHHmmssfff}.pdf";
                }
                // บังคับนามสกุลเป็น .pdf ถ้าไม่มี/ไม่ใช่
                if (!safeFileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                    safeFileName = Path.ChangeExtension(safeFileName, ".pdf");

                // 4) กันชื่อซ้ำ: ไม่ overwrite
                safeFileName = MakeUnique(targetFolder, safeFileName);

                var filePath = Path.Combine(targetFolder, safeFileName);

                // 5) บันทึกไฟล์
                try
                {
                    using var fs = new FileStream(filePath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
                    await file.CopyToAsync(fs);

                    // สร้าง URL สำหรับเรียกดู
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

            return Ok(new
            {
                message = "Upload completed",
                total = model.PostedFiles.Count,
                savedCount = saved.Count,
                skippedCount = skipped.Count,
                savedFiles = saved,
                skippedFiles = skipped
            });
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
