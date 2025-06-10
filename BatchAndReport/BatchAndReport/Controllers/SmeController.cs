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

        public SmeController(
            SmeDAO smeDao,
            IApiInformationRepository repositoryApi,
            ICallAPIService serviceApi)
        {
            _smeDao = smeDao;
            _repositoryApi = repositoryApi;
            _serviceApi = serviceApi;
        }

        [HttpPost("GetSME_Project")]
        public async Task<IActionResult> GetSME_Project([FromQuery] int page, [FromQuery] int perPage)
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
                        var projects = JsonSerializer.Deserialize<ApiResponseReturnProjectModels>(result.ToString());
                        if (projects?.Data == null) continue;

                        foreach (var item in projects.Data)
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
    }
}