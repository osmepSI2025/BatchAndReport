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
    public class SmeController : ControllerBase
    {
        private readonly SmeDAO _smeDao;
        private readonly IApiInformationRepository _repositoryApi;
        private readonly ICallAPIService _serviceApi;
        private readonly IPdfService _servicePdf;
        private readonly IWordService _serviceWord;

        public SmeController(
            SmeDAO smeDao,
            IApiInformationRepository repositoryApi,
            ICallAPIService serviceApi,
            IPdfService servicePdf,
            IWordService serviceWord)
        {
            _smeDao = smeDao;
            _repositoryApi = repositoryApi;
            _serviceApi = serviceApi;
            _servicePdf = servicePdf;
            _serviceWord = serviceWord;
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

        [HttpGet("ExportProjectDetailPdf")]
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

            var wordBytes = _serviceWord.GenerateWord(detail);
            var pdfBytes = _serviceWord.ConvertWordToPdf(wordBytes);
            return File(pdfBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                $"SME_Project_{projectCode}.pdf");
        }

        [HttpGet("ExportSMESummaryWord")]
        public async Task<IActionResult> ExportSMESummaryWord([FromQuery] string budYear)
        {
            var projects = await _smeDao.GetSummaryProjectAsync(budYear);
            var strategies = await _smeDao.GetProjectStrategyAsync(budYear);

            // ตรวจสอบว่ามีข้อมูลหรือไม่
            if (projects == null || !projects.Any())
                return NotFound("ไม่พบข้อมูลสำหรับปีงบประมาณที่ระบุ");

            var bytes = _serviceWord.GenerateSummaryWord(projects, strategies, budYear); // Pass 'budYear' as the second argument

            var pdfBytes = _serviceWord.ConvertWordToPdf(bytes);
            return File(
                pdfBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                $"SME_Summary_{budYear}.pdf"
            );
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

    }
}