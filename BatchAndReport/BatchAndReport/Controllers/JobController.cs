using BatchAndReport.DAO;
using BatchAndReport.Entities;
using BatchAndReport.Models;
using BatchAndReport.Repository;
using BatchAndReport.Services;
using DocumentFormat.OpenXml.InkML;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Text.Json;

namespace BatchAndReport.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class JobController : ControllerBase
    {
        private readonly ScheduledJobService _jobService;
        private readonly IApiInformationRepository _repositoryApi;
        private readonly ICallAPIService _serviceApi;
        private readonly HrDAO _hrDao;
        public JobController(ScheduledJobService jobService, IApiInformationRepository repositoryApi, ICallAPIService serviceApi, HrDAO hrDao)
        {
            _jobService = jobService;
            _repositoryApi = repositoryApi;
            _serviceApi = serviceApi;
            _hrDao = hrDao;
        }

        [HttpPost("run")]
        public async Task<IActionResult> RunJob([FromQuery] string jobName)
        {
            await _jobService.RunJobByNameAsync(jobName);
            return Ok(new { message = $"Job '{jobName}' triggered." });
        }


        [HttpPost("GetEmpHR")]
        public async Task<IActionResult> GetEmpHR([FromQuery] int page, int perPage)
        {
            try
            {
                var smodel = new searchEmployeeModels
                {
                    page = page,
                    perPage = perPage
                };

                var LApi = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "employee-all" });
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

                var result = await _serviceApi.GetDataApiAsync(apiParam, smodel);

                var employees = JsonSerializer.Deserialize<ApiListEmployeeResponse>(result.ToString());

                if (employees == null)
                    return BadRequest("Cannot deserialize employee data");

                var xdata = employees.Results.Select(emp => new MEmployeeModels
                {
                    EmployeeId = emp.EmployeeId,
                    EmployeeCode = emp.EmployeeCode,
                    NameTh = emp.NameTh,
                    NameEn = emp.NameEn,
                    FirstNameTh = emp.FirstNameTh,
                    FirstNameEn = emp.FirstNameEn,
                    LastNameTh = emp.LastNameTh,
                    LastNameEn = emp.LastNameEn,
                    Email = emp.Email,
                    Mobile = emp.Mobile,
                    EmploymentDate = emp.EmploymentDate,
                    TerminationDate = emp.TerminationDate,
                    EmployeeType = emp.EmployeeType,
                    EmployeeStatus = emp.EmployeeStatus,
                    SupervisorId = emp.SupervisorId,
                    CompanyId = emp.CompanyId,
                    BusinessUnitId = emp.BusinessUnitId,
                    PositionId = emp.PositionId
                }).ToList();

                await _hrDao.InsertOrUpdateEmployeesAsync(xdata);

                var supervisorApiInfo = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "employee-person" });
                var apiSup = supervisorApiInfo.Select(x => new MapiInformationModels
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

                if (apiSup == null) return BadRequest("Supervisor API info not found.");

                var supervisorData = new List<MEmployeeModels>();

                //var excludedIds = new HashSet<string>
                //{
                //    "R220817033714_SME_99999",
                //    "R221102113924_SME_62034",
                //    "R240130115907_SME_FS005"
                //};

                //xdata = xdata
                //    .Where(emp => !excludedIds.Contains(emp.EmployeeId))
                //    .ToList();

                foreach (var emp in xdata)
                {
                    if (string.IsNullOrEmpty(emp.EmployeeId))
                        continue;

                    // Rename the variable to avoid conflict
                    var supervisorResult = await _serviceApi.GetDataEmpMovementApiAsync(apiSup, emp.EmployeeId);

                    if (string.IsNullOrEmpty(supervisorResult))
                        continue;

                    try
                    {
                        var json = JsonSerializer.Deserialize<JsonElement>(supervisorResult);

                        string? jsonEmployeeId = null;
                        string? supervisorId = null;

                        if (json.TryGetProperty("results", out var resultsNode))
                        {
                            // ดึง employeeId จาก results
                            if (resultsNode.TryGetProperty("employeeId", out var empIdElement))
                            {
                                jsonEmployeeId = empIdElement.GetString();
                            }

                            // ดึง supervisorId จาก results
                            if (resultsNode.TryGetProperty("supervisorId", out var supElement))
                            {
                                supervisorId = supElement.GetString();
                            }
                        }

                        if (jsonEmployeeId?.Trim() == emp.EmployeeId.Trim() &&
                            !string.IsNullOrEmpty(supervisorId))
                        {
                            supervisorData.Add(new MEmployeeModels
                            {
                                EmployeeId = emp.EmployeeId,
                                SupervisorId = supervisorId
                            });
                        }
                        else
                        {
                            Console.WriteLine($"❌ Not matched or missing data: emp = [{emp.EmployeeId}], json = [{jsonEmployeeId}]");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error parsing JSON for {emp.EmployeeId}: {ex.Message}");
                    }


                }

                await _hrDao.UpdateSupervisorAsync(supervisorData);

                return Ok(new { message = "Sync and Save Complete", total = employees.Results.ToList().Count });
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

        [HttpPost("GetEmpMovement")]
        public async Task<IActionResult> GetEmpMovement([FromQuery] int page, int perPage)
        {
            try
            {
                // 🔍 1. โหลดข้อมูล Employee ทั้งหมดจาก API employee-all ก่อน
                var empListApi = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "employee-all" });
                var empApiParam = empListApi.Select(x => new MapiInformationModels
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

                if (empApiParam == null)
                    return BadRequest("Employee API info not found.");

                var empSearch = new searchEmployeeModels
                {
                    page = page,
                    perPage = perPage
                };

                var empRaw = await _serviceApi.GetDataApiAsync(empApiParam, empSearch);
                var empList = JsonSerializer.Deserialize<ApiListEmployeeResponse>(empRaw.ToString());

                if (empList?.Results == null || !empList.Results.Any())
                    return BadRequest("No employee data found.");

                // 🔁 2. เตรียมข้อมูล movement
                var movementListApi = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "employee-movement" });
                var movementApiParam = movementListApi.Select(x => new MapiInformationModels
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

                if (movementApiParam == null)
                    return BadRequest("Movement API info not found.");

                var allMovements = new List<MEmployeeMovementModels>();

                // 🔁 3. วนลูปจาก employee แต่ละคน แล้วเรียก API movement แยกตาม empId
                foreach (var emp in empList.Results)
                {
                    if (string.IsNullOrEmpty(emp?.EmployeeId))
                        continue;

                    var movementResult = await _serviceApi.GetDataEmpMovementApiAsync(movementApiParam, emp.EmployeeId);
                    var movementResponse = JsonSerializer.Deserialize<ApiListEmployeeMovementResponse>(movementResult.ToString());

                    if (movementResponse?.Results == null) continue;

                    var mapped = movementResponse.Results.Select(m => new MEmployeeMovementModels
                    {
                        EmployeeId = m.EmployeeId,
                        EffectiveDate = m.EffectiveDate,
                        MovementTypeId = m.MovementTypeId,
                        MovementReasonId = m.MovementReasonId,
                        EmployeeCode = m.EmployeeCode,
                        Employment = m.Employment,
                        EmployeeStatus = m.EmployeeStatus,
                        EmployeeTypeId = m.EmployeeTypeId,
                        PayrollGroupId = m.PayrollGroupId,
                        CompanyId = m.CompanyId,
                        BusinessUnitId = m.BusinessUnitId,
                        PositionId = m.PositionId,
                        WorkLocationId = m.WorkLocationId,
                        CalendarGroupId = m.CalendarGroupId,
                        JobTitleId = m.JobTitleId,
                        JobLevelId = m.JobLevelId,
                        JobGradeId = m.JobGradeId,
                        ContractStartDate = m.ContractStartDate,
                        ContractEndDate = m.ContractEndDate,
                        RenewContractCount = m.RenewContractCount,
                        ProbationDate = m.ProbationDate,
                        ProbationDuration = m.ProbationDuration,
                        ProbationResult = m.ProbationResult,
                        ProbationExtend = m.ProbationExtend,
                        EmploymentDate = m.EmploymentDate,
                        JoinDate = m.JoinDate,
                        OccupationDate = m.OccupationDate,
                        TerminationDate = m.TerminationDate,
                        TerminationReason = m.TerminationReason,
                        TerminationSSO = m.TerminationSSO,
                        IsBlacklist = m.IsBlacklist,
                        PaymentDate = m.PaymentDate,
                        Remark = m.Remark,
                        ServiceYearAdjust = m.ServiceYearAdjust,
                        SupervisorCode = m.SupervisorCode,
                        StandardWorkHoursID = m.StandardWorkHoursID,
                        WorkOperation = m.WorkOperation
                    }).ToList();

                    allMovements.AddRange(mapped);
                }

                await _hrDao.InsertOrUpdateEmployeesMovementAsync(allMovements);

                return Ok(new
                {
                    message = "Movement data sync complete.",
                    total = allMovements.Count
                });
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

        [HttpPost("GetPosition")]
        public async Task<IActionResult> GetPosition([FromQuery] int page, int perPage)
        {
            try
            {
                var smodel = new searchPositionModels
                {
                    page = page,
                    perPage = perPage
                };

                var LApi = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "position-all" });
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

                var result = await _serviceApi.GetDataApiAsync(apiParam, smodel);

                var positions = JsonSerializer.Deserialize<ApiListPositionResponse>(result.ToString());

                if (positions == null)
                    return BadRequest("Cannot deserialize position data");
                // Map PositionResult to MPositionModels
                var xdata = new List<MPositionModels>();
                xdata = positions.Results.Select(pos => new MPositionModels
                {
                    ProjectCode = pos.ProjectCode,
                    PositionId = pos.PositionId,
                    TypeCode = pos.TypeCode,
                    Module = pos.Module,
                    NameTh = pos.NameTh,
                    NameEn = pos.NameEn,
                    DescriptionTh = pos.DescriptionTh,
                    DescriptionEn = pos.DescriptionEn,
                }).ToList();
                await _hrDao.InsertOrUpdatePositionAsync(xdata);

                return Ok(new { message = "Sync and Save Complete", total = positions.Results.ToList().Count });
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

        [HttpPost("GetBusinessUnit")]
        public async Task<IActionResult> GetBusinessUnit([FromQuery] int page, int perPage)
        {
            try
            {
                var smodel = new SearchBusinessUnitModels
                {
                    page = page,
                    perPage = perPage
                };

                var LApi = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "business-units" });
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

                var result = await _serviceApi.GetDataApiAsync(apiParam, smodel);

                var businessUnits = JsonSerializer.Deserialize<ApiListBusinessUnitResponse>(result.ToString());

                if (businessUnits == null)
                    return BadRequest("Cannot deserialize business units data");
                // Map BusinessUnitsResult to MBusinessUnitsModels
                var xdata = new List<MBusinessUnitModels>();
                xdata = businessUnits.Results.Select(bus => new MBusinessUnitModels
                {
                    BusinessUnitId = bus.BusinessUnitId,
                    BusinessUnitCode = bus.BusinessUnitCode,
                    BusinessUnitLevel = bus.BusinessUnitLevel,
                    ParentId = bus.ParentId,
                    CompanyId = bus.CompanyId,
                    EffectiveDate = bus.EffectiveDate,
                    NameTh = bus.NameTh,
                    NameEn = bus.NameEn,
                    AbbreviationTh = bus.AbbreviationTh,
                    AbbreviationEn = bus.AbbreviationEn,
                    DescriptionTh = bus.DescriptionTh,
                    DescriptionEn = bus.DescriptionEn,
                    CreateDate = bus.CreateDate
                }).ToList();
                await _hrDao.InsertOrUpdateBusinessUnitAsync(xdata);

                return Ok(new { message = "Sync and Save Complete", total = businessUnits.Results.ToList().Count });
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
