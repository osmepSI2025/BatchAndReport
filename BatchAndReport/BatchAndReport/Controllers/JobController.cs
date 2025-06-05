using BatchAndReport.Entities;
using BatchAndReport.Models;
using BatchAndReport.Repository;
using BatchAndReport.Services;
using Microsoft.AspNetCore.Mvc;
using System.Text.Json;
using Microsoft.EntityFrameworkCore;

namespace BatchAndReport.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class JobController : ControllerBase
    {
        private readonly ScheduledJobService _jobService;
        private readonly IApiInformationRepository _repositoryApi;
        private readonly ICallAPIService _serviceApi;
        private readonly K2DBContext _k2context;
        public JobController(ScheduledJobService jobService, IApiInformationRepository repositoryApi, ICallAPIService serviceApi, K2DBContext k2context)
        {
            _jobService = jobService;
            _repositoryApi = repositoryApi;
            _serviceApi = serviceApi;
            _k2context = k2context;
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
                // Map EmployeeResult to MEmployeeModels
              var xdata = new List<MEmployeeModels>();
                xdata = employees.Results.Select(emp => new MEmployeeModels
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
                await InsertOrUpdateEmployeesAsync(xdata);

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

        public async Task InsertOrUpdateEmployeesAsync(List<MEmployeeModels> employees)
        {
            foreach (var emp in employees)
            {
                var existingEmp = await _k2context.Employees
                    .FirstOrDefaultAsync(e => e.EmployeeId == emp.EmployeeId);

                if (existingEmp != null)
                {
                    // UPDATE
                    existingEmp.EmployeeCode = emp.EmployeeCode;
                    existingEmp.NameTh = emp.NameTh;
                    existingEmp.NameEn = emp.NameEn;
                    existingEmp.FirstNameTh = emp.FirstNameTh;
                    existingEmp.FirstNameEn = emp.FirstNameEn;
                    existingEmp.LastNameTh = emp.LastNameTh;
                    existingEmp.LastNameEn = emp.LastNameEn;
                    existingEmp.Email = emp.Email;
                    existingEmp.Mobile = emp.Mobile;
                    existingEmp.EmploymentDate = emp.EmploymentDate;
                    existingEmp.TerminationDate = emp.TerminationDate;
                    existingEmp.EmployeeType = emp.EmployeeType;
                    existingEmp.EmployeeStatus = emp.EmployeeStatus;
                    existingEmp.SupervisorId = emp.SupervisorId;
                    existingEmp.CompanyId = emp.CompanyId;
                    existingEmp.BusinessUnitId = emp.BusinessUnitId;
                    existingEmp.PositionId = emp.PositionId;

                    _k2context.Employees.Update(existingEmp);
                }
                else
                {
                    // INSERT
                    var newEmp = new Employee
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
                    };

                    await _k2context.Employees.AddAsync(newEmp);
                }
            }

            await _k2context.SaveChangesAsync();
        }
    }
}