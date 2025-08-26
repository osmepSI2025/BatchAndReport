using BatchAndReport.DAO;
using BatchAndReport.Entities;
using BatchAndReport.Models;
using BatchAndReport.Repository;
using BatchAndReport.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.IO.Compression;
using System.Text;
using System.Text.Json;

namespace BatchAndReport.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class WorkflowController : ControllerBase
    {
        private readonly WorkflowDAO _workflowDao;
        private readonly IApiInformationRepository _repositoryApi;
        private readonly ICallAPIService _serviceApi;
        private readonly IPdfService _servicePdf;
        private readonly IWordWFService _serviceWFWord;
        private readonly WordWorkFlow_annualProcessReviewService _wordWorkFlow_AnnualProcessReviewService;
       private readonly WordSME_ReportService _ReportService;
        private readonly IConfiguration _configuration; // Add this line
        public WorkflowController(
            WorkflowDAO workflowDao,
            IApiInformationRepository repositoryApi,
            ICallAPIService serviceApi,
            IPdfService servicePdf,
            IWordWFService serviceWFWord,
            WordSME_ReportService reportService,
            WordWorkFlow_annualProcessReviewService wordWorkFlow_AnnualProcessReviewService,
            IConfiguration configuration)
        {
            _workflowDao = workflowDao;
            _repositoryApi = repositoryApi;
            _serviceApi = serviceApi;
            _servicePdf = servicePdf;
            _serviceWFWord = serviceWFWord;
            this._wordWorkFlow_AnnualProcessReviewService = wordWorkFlow_AnnualProcessReviewService;
            _ReportService = reportService;
            _configuration = configuration;
        }

        [HttpGet("ExportAnnualWorkProcesses")]
        public async Task<IActionResult> ExportAnnualWorkProcesses([FromQuery] int annualProcessReviewId)
        {
            var detail = await _workflowDao.GetProcessDetailAsync(annualProcessReviewId);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");     

            var pdfBytes = await _wordWorkFlow_AnnualProcessReviewService.GenAnnualWorkProcesses_HtmlToPDF(detail);
            return File(pdfBytes, "application/pdf", "AnnualWorkProcesses.pdf");
        }

        [HttpGet("ExportAnnualWorkProcessesWord")]
        public async Task<IActionResult> ExportAnnualWorkProcessesWord([FromQuery] int annualProcessReviewId)
        {
            var detail = await _workflowDao.GetProcessDetailAsync(annualProcessReviewId);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            var pdfBytes = await _wordWorkFlow_AnnualProcessReviewService.GenAnnualWorkProcesses_Html(detail);

            // 2. Convert HTML to Word document (byte array)
            var wordBytes = _ReportService.ConvertHtmlToWord(pdfBytes);

            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "AnnualWorkProcesses.docx");
        }

        [HttpGet("ExportAnnualWorkProcessesJPEG")]
        public async Task<IActionResult> ExportAnnualWorkProcessesJPEG([FromQuery] int annualProcessReviewId)
        {
            var detail = await _workflowDao.GetProcessDetailAsync(annualProcessReviewId);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            var pdfBytes = await _wordWorkFlow_AnnualProcessReviewService.GenAnnualWorkProcesses_HtmlToPDF(detail);

            // Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "WorkflowDocument", "AnnualWorkProcesses", "AnnualWorkProcesses_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"AnnualWorkProcesses.pdf");

            // Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // Convert each PDF page to JPEG
            var jpegPaths = new List<string>();
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"AnnualWorkProcesses_p{i + 1}.jpg");
                        using (var fs = new FileStream(jpegPath, FileMode.Create, FileAccess.Write))
                        {
                            image.Save(fs, System.Drawing.Imaging.ImageFormat.Jpeg);
                        }
                        jpegPaths.Add(jpegPath);
                    }
                }
            }
            // Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // Create zip file of JPEGs
            var folderPathzip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "WorkflowDocument", "AnnualWorkProcesses");
            var zipPath = Path.Combine(folderPathzip, $"AnnualWorkProcesses_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegPaths)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // Delete JPEG files after zipping
            foreach (var file in jpegPaths)
            {
                System.IO.File.Delete(file);
            }

            // Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"AnnualWorkProcesses_JPEG.zip");
        }


        [HttpGet("ExportWorkSystem")]
        public async Task<IActionResult> ExportWorkSystem(
            [FromQuery] int? fiscalYear = null,
            [FromQuery] string? businessUnitId = null,
            [FromQuery] string? processTypeCode = null,
            [FromQuery] string? processGroupCode = null,
            [FromQuery] string? processCode = null,
            [FromQuery] int? processCategory = null) // Changed type from int? to string?
        {
            var detail = await _workflowDao.GetWorkSystemDataAsync(
                fiscalYear,
                businessUnitId,
                processTypeCode,
                processGroupCode,
                processCode,
                processCategory // Updated to match the expected type in the method signature
            );

            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            var generator = _serviceWFWord.GenWorkSystem(detail);
            var excelBytes = generator; // Assuming `GenWorkSystem()` already returns a byte array.

            return File(excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "AnnualWorkProcesses.xlsx");
        }

        [HttpGet("ExportInternalControl")]
        public async Task<IActionResult> ExportInternalControl(
            [FromQuery] int? fiscalYear = null,
            [FromQuery] string? businessUnitId = null,
            [FromQuery] string? processTypeCode = null,
            [FromQuery] string? processGroupCode = null,
            [FromQuery] string? processCode = null,
            [FromQuery] int? processCategory = null)
        {
            var detail = await _workflowDao.GetInternalControlProcessesAsync(fiscalYear,
                businessUnitId,
                processTypeCode,
                processGroupCode,
                processCode,
                processCategory);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            var generator = _serviceWFWord.GenInternalControlSystem(detail);
            var excelBytes = generator; // Assuming `GenWorkSystem()` already returns a byte array.

            return File(excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "InternalControl.xlsx");
        }

        [HttpGet("ExportInternalControlEncypt")]
        public async Task<IActionResult> ExportInternalControlEncypt(
       [FromQuery] int? fiscalYear = null,
       [FromQuery] string? businessUnitId = null,
       [FromQuery] string? processTypeCode = null,
       [FromQuery] string? processGroupCode = null,
       [FromQuery] string? processCode = null,
       [FromQuery] int? processCategory = null)
        {
            var detail = await _workflowDao.GetInternalControlProcessesAsync(fiscalYear,
                businessUnitId,
                processTypeCode,
                processGroupCode,
                processCode,
                processCategory);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            var generator = _serviceWFWord.GenInternalControlSystem(detail);
            var excelBytes = generator; // Assuming this is a valid Excel file (byte[])

            // Get password from appsettings.json
            var password = _configuration["Password:PaswordPDF"];

            if (!string.IsNullOrEmpty(password))
            {
                using var inputStream = new MemoryStream(excelBytes);
                using var package = new OfficeOpenXml.ExcelPackage(inputStream);
                package.Encryption.IsEncrypted = true;
                package.Encryption.Password = password;

                using var outputStream = new MemoryStream();
                package.SaveAs(outputStream);
                outputStream.Position = 0;
                return File(outputStream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "InternalControl_ProtectedEncypt.xlsx");
            }
            else
            {
                return File(excelBytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "InternalControlEncypt.xlsx");
            }
        }


        [HttpGet("ExportInternalControlDoc")]
        public async Task<IActionResult> ExportInternalControlDoc(
            [FromQuery] int? subProcessId = null,
            [FromQuery] int? processId = null)
        {
            if (processId == null)
                return BadRequest("ProcessId is required.");

            var detail = await _workflowDao.GetInternalControlProcessesByProcessID(processId.Value);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            if (subProcessId == null)
                return BadRequest("SubProcessId is required.");

            var detail2 = await _workflowDao.GetSubProcessDetailAsync(subProcessId.Value); // Use .Value to pass int instead of int?
            if (detail2 == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            var wordBytes = await _serviceWFWord.GenInternalControlSystemWord(detail, detail2);
            var pdfBytes = _serviceWFWord.ConvertWordToPdf(wordBytes);
            return File(wordBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                $"Internal_WorkProcessPoint.docx");
        }

        [HttpGet("ExportWorkProcessPoint")]
        public async Task<IActionResult> ExportWorkProcessPoint([FromQuery] int subProcessId)
        {
            var detail = await _workflowDao.GetSubProcessDetailAsync(subProcessId);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");


            var wordBytes = await _serviceWFWord.GenWorkProcessPointHtmlToPdf(detail);
            return File(wordBytes, "application/pdf", "WorkProcessPoint.pdf");
        }

        [HttpGet("ExportWorkflowProcess")]
        public async Task<IActionResult> ExportWorkflowProcess([FromQuery] int idParam)
        {
            var detail = await _workflowDao.GetWFProcessDetailAsync(idParam);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

           // var wordBytes = _serviceWFWord.GenWFProcessDetail(detail);
            //var pdfBytes = _serviceWFWord.ConvertWordToPdf(wordBytes);
            var pdfBytes = await _wordWorkFlow_AnnualProcessReviewService.GenExportWorkProcesses_HtmlToPDF(detail);
            return File(pdfBytes, "application/pdf", "ExportWorkProcesses.pdf");
        }
        [HttpGet("ExportWorkflowProcessTXT")]
        public async Task<IActionResult> ExportWorkflowProcessTXT([FromQuery] int idParam)
        {
            var detail = await _workflowDao.GetWFProcessDetailAsync(idParam);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("=== Core Processes ===");
            if (detail.CoreProcesses != null)
            {
                foreach (var item in detail.CoreProcesses)
                {
                    sb.AppendLine($"- {item.ProcessGroupCode+":"+ item.ProcessGroupName}");
                }
            }
            sb.AppendLine();
            sb.AppendLine("=== Support Processes ===");
            if (detail.SupportProcesses != null)
            {
                foreach (var item in detail.SupportProcesses)
                {
                    sb.AppendLine($"- {item.ProcessGroupCode + ":" + item.ProcessGroupName}");
                }
            }

            var bytes = System.Text.Encoding.UTF8.GetBytes(sb.ToString());
            return File(bytes, "text/plain", "ExportWorkProcesses.txt");
        }
        [HttpGet("ExportWorkflowProcessJPEG")]
        public async Task<IActionResult> ExportWorkflowProcessJPEG([FromQuery] int idParam)
        {
            var detail = await _workflowDao.GetWFProcessDetailAsync(idParam);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            // 1. Generate HTML from the detail
            var pdfBytes = await _wordWorkFlow_AnnualProcessReviewService.GenExportWorkProcesses_HtmlToPDF(detail);
            // Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "WorkflowDocument", "WorkProcesses", "WorkProcesses_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"WorkProcesses.pdf");

            // Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // Convert each PDF page to JPEG
            var jpegPaths = new List<string>();
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"ExportWorkProcesses_p{i + 1}.jpg");
                        using (var fs = new FileStream(jpegPath, FileMode.Create, FileAccess.Write))
                        {
                            image.Save(fs, System.Drawing.Imaging.ImageFormat.Jpeg);
                        }
                        jpegPaths.Add(jpegPath);
                    }
                }
            }
            // Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // Create zip file of JPEGs
            var folderPathzip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "WorkflowDocument", "WorkProcesses");
            var zipPath = Path.Combine(folderPathzip, $"WorkProcesses_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegPaths)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // Delete JPEG files after zipping
            foreach (var file in jpegPaths)
            {
                System.IO.File.Delete(file);
            }

            // Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"WorkProcesses_JPEG.zip");
        }


        [HttpGet("ExportWorkflowProcessWord")]
        public async Task<IActionResult> ExportWorkflowProcessWord([FromQuery] int idParam)
        {
            var detail = await _workflowDao.GetWFProcessDetailAsync(idParam);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

     
            var pdfBytes = await _wordWorkFlow_AnnualProcessReviewService.GenExportWorkProcesses_Html(detail);
            // Convert HTML to Word document (byte array)
            var wordBytes = _ReportService.ConvertHtmlToWord(pdfBytes);

            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "ExportWorkProcesses.docx");
        
        }
        [HttpGet("ExportCreateWFStatus")]
        public async Task<IActionResult> ExportCreateWFStatus(
            [FromQuery] int? fiscalYearId = null,
            [FromQuery] string? businessUnitId = null,
            [FromQuery] string? processTypeCode = null,
            [FromQuery] string? processGroupCode = null,
            [FromQuery] string? processCode = null,
            [FromQuery] int? processCategory = null,
            [FromQuery] bool? isST01 = null,
            [FromQuery] bool? isST0101 = null,
            [FromQuery] bool? isST0102 = null,
            [FromQuery] bool? isST0103 = null,
            [FromQuery] bool? isST0104 = null,
            [FromQuery] bool? isST0105 = null
        )
        {
            var detail = await _workflowDao.GetCreateProcessStatusAsync(
                fiscalYearId,
                businessUnitId,
                processTypeCode,
                processGroupCode,
                processCode,
                isST01,
                isST0101,
                isST0102,
                isST0103,
                isST0104,
                isST0105

            );

            if (detail == null)
                return NotFound("ไม่พบข้อมูล");

            var generator = _serviceWFWord.GenCreateWFStatus(detail);
            var excelBytes = generator; // Assuming `GenWorkSystem()` already returns a byte array.

            return File(excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "CreateWFStatus.xlsx");
        }
        [HttpGet("ExportProcessResultByIndicator")]
        public async Task<IActionResult> ExportProcessResultByIndicator(
            [FromQuery] int? fiscalYearId = null,
            [FromQuery] string? businessUnitId = null,
            [FromQuery] string? processTypeCode = null,
            [FromQuery] string? processCode = null,
            [FromQuery] bool? isEvaluationTrue = null,
            [FromQuery] bool? isEvaluationFalse = null,
            [FromQuery] int? subMasterProcessId = null

        )
        // Changed type from int? to string?
        {
            var detail = await _workflowDao.GetProcessResultByIndicatorAsync(
                fiscalYearId,
                businessUnitId,
                processTypeCode,
                processCode,
                isEvaluationTrue,
                isEvaluationFalse,
                subMasterProcessId
            );

            if (detail == null)
                return NotFound("ไม่พบข้อมูล");

            var generator = _serviceWFWord.GenProcessResultByIndicator(detail);
            var excelBytes = generator;

            return File(excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "ProcessResultByIndicator.xlsx");
        }

        [HttpGet("Workflow")]
        [Produces("application/json")]
        public async Task<IActionResult> GetWorkflow()
        {
            string json;
            try
            {
                json = await _workflowDao.GetSubProcessMaterAsync();
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

        [HttpGet("WorkflowActivity")]
        [Produces("application/json")]
        public async Task<IActionResult> GetWorkflowActivity()
        {
            string json;
            try
            {
                json = await _workflowDao.GetWorkflowActivityAsync();
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