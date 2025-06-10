using BatchAndReport.DAO;
using BatchAndReport.Models;
using BatchAndReport.Repository;
using BatchAndReport.Services;
using Microsoft.AspNetCore.Mvc;
using System.Text.Json;

namespace BatchAndReport.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class EContractController : ControllerBase
    {
        private readonly EContractDAO _eContractDao;
        private readonly IApiInformationRepository _repositoryApi;
        private readonly ICallAPIService _serviceApi;

        public EContractController(
            EContractDAO eContractDao,
            IApiInformationRepository repositoryApi,
            ICallAPIService serviceApi)
        {
            _eContractDao = eContractDao;
            _repositoryApi = repositoryApi;
            _serviceApi = serviceApi;
        }

        [HttpPost("GetEmpContract")]
        public async Task<IActionResult> GetEmpContract([FromQuery] string employmentDate)
        {
            try
            {
                var apiList = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "employee-contract" });
                var apiParam = apiList.Select(x => new MapiInformationModels
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

                // Prepare filter model
                var filterModel = new
                {
                    employmentDate = employmentDate
                };

                var result = await _serviceApi.GetDataApiAsync(apiParam, filterModel);

                var employees = JsonSerializer.Deserialize<ApiListEmployeeContractResponse>(result.ToString());

                if (employees?.Results == null || !employees.Results.Any())
                    return BadRequest("No employee contract data found.");

                // Map to model
                var contractModels = employees.Results
                    .Where(emp => emp != null) // Ensure no null references
                    .Select(emp => new MEmployeeContractModels
                    {
                        ContractFlag = emp!.ContractFlag, // Use null-forgiving operator
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
                        PositionId = emp.PositionId,
                        Salary = emp.Salary,
                        IdCard = emp.IdCard,
                        PassportNo = emp.PassportNo
                    }).ToList();

                await _eContractDao.InsertOrUpdateEmployeeContractsAsync(contractModels);

                return Ok(new { message = "Sync and Save Complete", total = contractModels.Count });
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