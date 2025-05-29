using BatchAndReport.Models;
using BatchAndReport.Repository;
using BatchAndReport.Services;
using Microsoft.AspNetCore.Mvc;

namespace BatchAndReport.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class JobController : ControllerBase
    {
        private readonly ScheduledJobService _jobService;
        private readonly IApiInformationRepository _repositoryApi;
        private readonly ICallAPIService _serviceApi;
        public JobController(ScheduledJobService jobService, IApiInformationRepository repositoryApi, ICallAPIService serviceApi)
        {
            _jobService = jobService;
            _repositoryApi = repositoryApi;
            _serviceApi = serviceApi;
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
            }).First(); // ดึงตัวแรกของ List


            var result = await _serviceApi.GetDataApiAsync(apiParam, smodel);
            return Ok(result);
        }
    }
}