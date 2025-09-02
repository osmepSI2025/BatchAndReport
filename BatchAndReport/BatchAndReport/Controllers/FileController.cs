// Controllers/FileController.cs
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using BatchAndReport.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace BatchAndReport.Controllers
{
    public class FileController : Controller
    {
        private readonly IWebHostEnvironment _env;
        public FileController(IWebHostEnvironment env) => _env = env;

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Upload(FileUploadModel model)
        {
            if (model.PostedFiles == null || !model.PostedFiles.Any())
                return BadRequest("No files uploaded.");

            // โฟลเดอร์ปลายทาง – เลือกได้ว่าต้องการแบบไหน:
            // var targetFolder = Path.Combine(_env.WebRootPath, "Document", "K2");
            // หรือให้แยกตาม ProcessInstanceID (ถ้าส่งมา):
            var targetFolder = string.IsNullOrWhiteSpace(model.ProcessInstanceID)
                ? Path.Combine(_env.WebRootPath, "Document", "ImportContract")
                : Path.Combine(_env.WebRootPath, "Document", "ImportContract", model.ProcessInstanceID);

            Directory.CreateDirectory(targetFolder);

            foreach (var file in model.PostedFiles)
            {
                if (file.Length <= 0) continue;

                // กัน path traversal: เอาเฉพาะชื่อไฟล์จริง
                var safeFileName = Path.GetFileName(file.FileName);

                var filePath = Path.Combine(targetFolder, safeFileName);

                // เขียนทับไฟล์เดิมถ้ามีชื่อซ้ำ (ไม่ rename)
                using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
                await file.CopyToAsync(stream);
            }

            // กลับไปหน้าเดิมพร้อมข้อความ หรือจะคืน JSON ก็ได้
            // return RedirectToAction("Index"); // ถ้ามีหน้าแสดงผล
            return Ok(new { message = "Upload complete.", folder = targetFolder, count = model.PostedFiles.Count });
        }
    }
}
