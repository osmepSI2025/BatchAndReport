using BatchAndReport.DAO;
using BatchAndReport.Entities;
using BatchAndReport.Models;
using BatchAndReport.Repository;
using BatchAndReport.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Syncfusion.Pdf.Graphics;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.IO.Compression;
using System.Text;
using System.Text.Json;

namespace BatchAndReport.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class SmeController : ControllerBase
    {
        private readonly SmeDAO _smeDao;
        private readonly IApiInformationRepository _repositoryApi;
        private readonly ICallAPIService _serviceApi;
        private readonly IPdfService _servicePdf;
        private readonly IWordService _serviceWord;
        private readonly WordSME_ReportService _wordSME_ReportService;

        public SmeController(
            SmeDAO smeDao,
            IApiInformationRepository repositoryApi,
            ICallAPIService serviceApi,
            IPdfService servicePdf,
            IWordService serviceWord,
            WordSME_ReportService wordSME_ReportService)
        {
            _smeDao = smeDao;
            _repositoryApi = repositoryApi;
            _serviceApi = serviceApi;
            _servicePdf = servicePdf;
            _serviceWord = serviceWord;
            _wordSME_ReportService = wordSME_ReportService;
        }

        [HttpGet("GetSME_Project")]
        public async Task<IActionResult> GetSME_Project()
        {
            try
            {
                var fiscalYears = await _smeDao.GetDistinctFiscalYearsAsync();

                var LApi = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "Return_Project" });
                var apiParam = LApi.Select(x => new MapiInformationModels
                {
                    ServiceNameCode = x.ServiceNameCode,
                    ApiKey = x.ApiKey,
                    AuthorizationType = x.AuthorizationType,
                    ContentType = x.ContentType,
                    CreateDate = x.CreateDate,
                    Id = x.Id,
                    MethodType = x.MethodType,
                    ServiceNameTh = x.ServiceNameTh,
                    Urldevelopment = x.Urldevelopment,
                    Urlproduction = x.Urlproduction,
                    Username = x.Username,
                    Password = x.Password,
                    UpdateDate = x.UpdateDate
                }).FirstOrDefault();

                if (apiParam == null)
                    return BadRequest("API info not found.");

                var allProjectData = new List<MProjectMasterModels>();

                foreach (var year in fiscalYears)
                {
                    var result = await _serviceApi.GetDataByParamApiAsync(apiParam, year);
                    if (string.IsNullOrEmpty(result))
                    {
                        return BadRequest($"API returned null or empty for year: {year}");
                    }
                    try
                    {
                        var projects = JsonSerializer.Deserialize<Dictionary<string, ProjectMasterResult>>(result);
                        if (projects == null) continue;

                        foreach (var item in projects)
                        {
                            if (!int.TryParse(item.Key, out int keyId))
                                continue;

                            var newProject = new MProjectMasterModels
                            {
                                KeyId = keyId,
                                ProjectName = item.Value.DATA_P1,
                                BudgetAmount = item.Value.DATA_P12,
                                Issue = item.Value.DATA_P6,
                                Strategy = item.Value.DATA_P9,
                                FiscalYear = year
                            };

                            allProjectData.Add(newProject);
                        }
                    }
                    catch (JsonException jex)
                    {
                        return StatusCode(500, new
                        {
                            message = $"JSON parse error for year: {year}",
                            jsonError = jex.Message,
                            rawResult = result
                        });
                    }
                }

                await _smeDao.InsertOrUpdateProjectMasterAsync(allProjectData);

                return Ok(new { message = "Sync and Save Complete", total = allProjectData.Count });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    message = "Internal Server Error",
                    error = ex.Message,
                    inner = ex.InnerException?.Message,
                    stack = ex.StackTrace
                });
            }
        }

        [HttpGet("ExportProjectDetailPdf_old")]
        public async Task<IActionResult> ExportPdf([FromQuery] string projectCode)
        {
            var detail = await _smeDao.GetProjectDetailAsync(projectCode);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            var pdfBytes = _servicePdf.GeneratePdf(detail);
            return File(pdfBytes, "application/pdf", $"SME_Project_{projectCode}.pdf");
        }

        [HttpGet("ExportProjectDetailWord")]
        public async Task<IActionResult> ExportProjectDetailWord([FromQuery] string projectCode)
        {
            var detail = await _smeDao.GetProjectDetailAsync(projectCode);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            // Option 1: If you want to generate Word from HTML (recommended for rich formatting)
            var html = await _wordSME_ReportService.ExportSMEProjectDetail_HTML(detail, projectCode);
            var wordBytes = _wordSME_ReportService.ConvertHtmlToWord(html);

            // Option 2: If your _serviceWord.GenerateWord(detail) already works, keep using it
            // var wordBytes = _serviceWord.GenerateWord(detail);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SMEDocument", "SME_Detail", "SME_" + projectCode);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"SME_{projectCode}.docx");

            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"SME_{projectCode}.docx");
        }


        [HttpGet("ExportProjectDetailPDF")]
        public async Task<IActionResult> ExportProjectDetailPDF([FromQuery] string projectCode)
        {
            var detail = await _smeDao.GetProjectDetailAsync(projectCode);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            //var wordBytes = _serviceWord.GenerateWord(detail);
            //var pdfBytes = _serviceWord.ConvertWordToPdf(wordBytes);
            //return File(pdfBytes,
            //    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            //    $"SME_Project_{projectCode}.pdf");
            var pdfBytes = await _wordSME_ReportService.ExportSMEProjectDetail_ToPDF(detail, projectCode);

            return File(
                pdfBytes,
                "application/pdf",
                 $"SME_Project_{projectCode}.pdf"
            );
        }
        [HttpGet("ExportProjectDetailJPEG")]
        public async Task<IActionResult> ExportProjectDetailJPEG([FromQuery] string projectCode)
        {
            var detail = await _smeDao.GetProjectDetailAsync(projectCode);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            // Generate PDF first
            var pdfBytes = await _wordSME_ReportService.ExportSMEProjectDetail_ToPDF(detail, projectCode);

            // Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SMEDocument", "Detail", $"SME_{projectCode}_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"SME_{projectCode}.pdf");

            // Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"SME_{projectCode}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // Create zip file of JPEGs
            var zipPath = Path.Combine(folderPath, $"SME_{projectCode}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"SME_{projectCode}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"SME_{projectCode}_JPEG.zip");
        }
        [HttpGet("ExportSMESummaryWord")]
        public async Task<IActionResult> ExportSMESummaryWord([FromQuery] string budYear)
        {
            var projects = await _smeDao.GetSummaryProjectAsync(budYear);
            var strategies = await _smeDao.GetProjectStrategyAsync(budYear);

            // ตรวจสอบว่ามีข้อมูลหรือไม่
            if (projects == null || !projects.Any())
                return NotFound("ไม่พบข้อมูลสำหรับปีงบประมาณที่ระบุ");

            // Generate Word document
            var wordBytes = _serviceWord.GenerateSummaryWord(projects, strategies, budYear);

            return File(
                wordBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                $"SME_Summary_{budYear}.docx"
            );
        }

        [HttpGet("ExportSMESummaryPDF")]
        public async Task<IActionResult> ExportSMESummaryPDF([FromQuery] string budYear)
        {
            var projects = await _smeDao.GetSummaryProjectAsync(budYear);
            var strategies = await _smeDao.GetProjectStrategyAsync(budYear);

            // var xdata = strategies.Distinct().ToList();
            // ตรวจสอบว่ามีข้อมูลหรือไม่
            if (projects == null || !projects.Any())
                return NotFound("ไม่พบข้อมูลสำหรับปีงบประมาณที่ระบุ");

     
            var pdfBytes = await _wordSME_ReportService.GenerateSummarySME_Budget_ToPdf(projects, strategies, budYear);
            return File(
                pdfBytes,
                "application/pdf",
                $"SME_Summary_{budYear}.pdf"
            );
        }

        [HttpGet("ExportSMESummaryJPEG")]
        public async Task<IActionResult> ExportSMESummaryJPEG([FromQuery] string budYear)
        {
            var projects = await _smeDao.GetSummaryProjectAsync(budYear);
            var strategies = await _smeDao.GetProjectStrategyAsync(budYear);
            // ตรวจสอบว่ามีข้อมูลหรือไม่
            if (projects == null || !projects.Any())
                return NotFound("ไม่พบข้อมูลสำหรับปีงบประมาณที่ระบุ");

            // You need to implement this method in your service
            var jpegBytes = await _wordSME_ReportService.GenerateSummarySME_Budget_ToPdf(projects, strategies, budYear);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SMEDocument", "Summary", "SME_" + budYear, "SME_" + budYear + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"SME_{budYear}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, jpegBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(jpegBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"EC_{budYear}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SMEDocument", "Summary", "SME_" + budYear);
            var zipPath = Path.Combine(folderPathZip, $"SME_{budYear}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"SME_{budYear}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"SME_{budYear}_JPEG.zip");
        }

        [HttpPost("SyncFiscalYears")]
        public async Task<IActionResult> SyncFiscalYears()
        {
            int currentYear = DateTime.Now.Year;
            int range = 5;

            // แปลง ค.ศ. เป็น พ.ศ.
            var fiscalYears = Enumerable
                .Range(currentYear - range, (range * 2) + 1)
                .Select(y => y + 543)
                .ToList();

            await _smeDao.InsertOrUpdateFiscalYearsAsync(fiscalYears);

            return Ok(new { message = "Sync Complete", years = fiscalYears });
        }

        // กลยุทธ์ strategy sme
        [HttpGet("strategy")]
        public async Task<IActionResult> strategy([FromQuery] int year)
        {
            var details = await _smeDao.GetStrategyDetailsByYearAsync(year);

            // Return the details directly (StrategyResponse already contains responseCode, etc.)
            return Ok(details);
        }

        [HttpGet("SME_Project")]
        [Produces("application/json")]
        public async Task<IActionResult> GetSmeProject([FromQuery] string year)
        {
            if (string.IsNullOrWhiteSpace(year))
            {
                return BadRequest(new
                {
                    responseCode = "400",
                    responseMsg = "Missing required query param 'year'",
                    data = Array.Empty<object>()
                });
            }

            string json;
            try
            {
                json = await _smeDao.GetSmeProjectFlatByYearAsync(year);
            }
            catch (Exception ex)
            {
                var err = new { responseCode = "500", responseMsg = "Database error: " + ex.Message, data = Array.Empty<object>() };
                return Content(System.Text.Json.JsonSerializer.Serialize(err), "application/json", Encoding.UTF8);
            }

            // sanity check แบบเบาๆ ว่าสตริงหน้าตาเป็น JSON
            if (string.IsNullOrWhiteSpace(json) || !(json.TrimStart().StartsWith("{") || json.TrimStart().StartsWith("[")))
            {
                var err = new { responseCode = "500", responseMsg = "Stored procedure returned invalid JSON.", data = Array.Empty<object>() };
                return Content(System.Text.Json.JsonSerializer.Serialize(err), "application/json", Encoding.UTF8);
            }

            // สำคัญ: ส่งเป็น application/json ตรง ๆ — ไม่ Ok(string) (จะถูก escape)
            return Content(json, "application/json", Encoding.UTF8);
        }
    }
}