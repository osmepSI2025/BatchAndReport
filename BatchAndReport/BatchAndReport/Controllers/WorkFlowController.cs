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
        public async Task<IActionResult> ExportAnnualWorkProcesses()
        {
            //var detail = await _workflowDao.GetProjectDetailAsync(projectCode);
            //if (detail == null)
            //    return NotFound("ไม่พบข้อมูลโครงการ");

            var wordBytes = _serviceWFWord.GenAnnualWorkProcesses();
            var pdfBytes = _serviceWFWord.ConvertWordToPdf(wordBytes);
            return File(pdfBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                $"AnnualWorkProcesses_test.pdf");
        }

        [HttpGet("ExportWorkSystem")]
        public async Task<IActionResult> ExportWorkSystem()
        {
            //var detail = await _workflowDao.GetProjectDetailAsync(projectCode);
            //if (detail == null)
            //    return NotFound("ไม่พบข้อมูลโครงการ");

            var generator = _serviceWFWord.GenWorkSystem();
            var excelBytes = generator; // Assuming `GenWorkSystem()` already returns a byte array.

            return File(excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "AnnualWorkProcesses.xlsx");
        }

    }
}