using Microsoft.AspNetCore.Mvc;

namespace BatchAndReport.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class JobController : ControllerBase
    {
        private readonly ScheduledJobService _jobService;

        public JobController(ScheduledJobService jobService)
        {
            _jobService = jobService;
        }

        [HttpPost("run")]
        public async Task<IActionResult> RunJob([FromQuery] string jobName)
        {
            await _jobService.RunJobByNameAsync(jobName);
            return Ok(new { message = $"Job '{jobName}' triggered." });
        }
    }
}