using BatchAndReport.Entities;
using Microsoft.EntityFrameworkCore;

public class ScheduledJobService : BackgroundService
{
    private readonly IServiceProvider _serviceProvider;
    private readonly ILogger<ScheduledJobService> _logger;

    public ScheduledJobService(IServiceProvider serviceProvider, ILogger<ScheduledJobService> logger)
    {
        _serviceProvider = serviceProvider;
        _logger = logger;
    }

    public async Task RunJobByNameAsync(string jobName)
    {
        using (var scope = _serviceProvider.CreateScope())
        {
            var db = scope.ServiceProvider.GetRequiredService<BatchDBContext>();
            var job = await db.MscheduledJobs
                .FirstOrDefaultAsync(j => j.IsActive == true && j.JobName == jobName);

            if (job != null)
            {
                try
                {
                    // TODO: Replace with your actual job logic
                    _logger.LogInformation($"Manually running job: {job.JobName} at {DateTime.Now}");
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"Error running job: {job.JobName}");
                }
            }
            else
            {
                _logger.LogWarning($"Job '{jobName}' not found or inactive.");
            }
        }
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            using (var scope = _serviceProvider.CreateScope())
            {
                var db = scope.ServiceProvider.GetRequiredService<BatchDBContext>();
                var now = DateTime.Now;
                var jobs = await db.MscheduledJobs
                    .Where(j => j.IsActive == true && j.RunHour == now.Hour && j.RunMinute == now.Minute)
                    .ToListAsync(stoppingToken);

                foreach (var job in jobs)
                {
                    try

                    {
                        //switchcase (job.JobName)
                        //{
                        //
                        //}

                        // TODO: Replace with your actual job logic
                        _logger.LogInformation($"Running job: {job.JobName} at {now}");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, $"Error running job: {job.JobName}");
                    }
                }
            }

            // Wait until the next minute
            var delay = 60000 - (DateTime.Now.Second * 1000 + DateTime.Now.Millisecond);
            await Task.Delay(delay > 0 ? delay : 1000, stoppingToken);
        
        }
    }


}