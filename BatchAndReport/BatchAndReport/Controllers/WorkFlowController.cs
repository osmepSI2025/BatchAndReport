using BatchAndReport.DAO;
using BatchAndReport.Entities;
using BatchAndReport.Models;
using BatchAndReport.Repository;
using BatchAndReport.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
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

        public WorkflowController(
            WorkflowDAO workflowDao,
            IApiInformationRepository repositoryApi,
            ICallAPIService serviceApi,
            IPdfService servicePdf,
            IWordWFService serviceWFWord)
        {
            _workflowDao = workflowDao;
            _repositoryApi = repositoryApi;
            _serviceApi = serviceApi;
            _servicePdf = servicePdf;
            _serviceWFWord = serviceWFWord;
        }

        [HttpGet("ExportAnnualWorkProcesses")]
        public async Task<IActionResult> ExportAnnualWorkProcesses([FromQuery] int fiscalYear)
        {
            var detail = await _workflowDao.GetProcessDetailAsync(fiscalYear);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            var wordBytes = _serviceWFWord.GenAnnualWorkProcesses(detail);
            var pdfBytes = _serviceWFWord.ConvertWordToPdf(wordBytes);
            return File(pdfBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                $"AnnualWorkProcesses_test.pdf");
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
        public async Task<IActionResult> ExportInternalControl([FromQuery] int processID)
        {
            var detail = await _workflowDao.GetInternalControlProcessesAsync(processID);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            var generator = _serviceWFWord.GenInternalControlSystem(detail);
            var excelBytes = generator; // Assuming `GenWorkSystem()` already returns a byte array.

            return File(excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "InternalControl.xlsx");
        }

        [HttpGet("ExportWorkProcessPoint")]
        public async Task<IActionResult> ExportWorkProcessPoint([FromQuery] int subProcessId)
        {
            var detail = await _workflowDao.GetSubProcessDetailAsync(subProcessId);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            var wordBytes = await _serviceWFWord.GenWorkProcessPoint(detail);
            var pdfBytes = _serviceWFWord.ConvertWordToPdf(wordBytes);
            return File(pdfBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                $"WorkProcessPoint_test.pdf");
        }

        [HttpGet("ExportWorkflowProcess")]
        public async Task<IActionResult> ExportWorkflowProcess([FromQuery] int idParam)
        {
            var detail = await _workflowDao.GetWFProcessDetailAsync(idParam);
            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            var wordBytes = _serviceWFWord.GenWFProcessDetail(detail);
            var pdfBytes = _serviceWFWord.ConvertWordToPdf(wordBytes);
            return File(pdfBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                $"WFProcessDetail.pdf");
        }

    }
}