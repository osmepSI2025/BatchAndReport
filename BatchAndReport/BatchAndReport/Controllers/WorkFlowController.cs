using BatchAndReport.DAO;
using BatchAndReport.Entities;
using BatchAndReport.Models;
using BatchAndReport.Repository;
using BatchAndReport.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
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
        public WorkflowController(
            WorkflowDAO workflowDao,
            IApiInformationRepository repositoryApi,
            ICallAPIService serviceApi,
            IPdfService servicePdf,
            IWordWFService serviceWFWord,
            WordWorkFlow_annualProcessReviewService wordWorkFlow_AnnualProcessReviewService)
        {
            _workflowDao = workflowDao;
            _repositoryApi = repositoryApi;
            _serviceApi = serviceApi;
            _servicePdf = servicePdf;
            _serviceWFWord = serviceWFWord;
            this._wordWorkFlow_AnnualProcessReviewService = wordWorkFlow_AnnualProcessReviewService;
        }

        [HttpGet("ExportAnnualWorkProcesses")]
        public async Task<IActionResult> ExportAnnualWorkProcesses([FromQuery] int annualProcessReviewId)
        {
            var detail = await _workflowDao.GetProcessDetailAsync(annualProcessReviewId);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            //var wordBytes = _serviceWFWord.GenAnnualWorkProcesses(detail);
            //var pdfBytes = _serviceWFWord.ConvertWordToPdf(wordBytes);
            //return File(pdfBytes,
            //    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            //    $"AnnualWorkProcesses.pdf");

            var pdfBytes = await _wordWorkFlow_AnnualProcessReviewService.GenAnnualWorkProcesses_HtmlToPDF(detail);
            return File(pdfBytes, "application/pdf", "AnnualWorkProcesses.pdf");
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

    }
}