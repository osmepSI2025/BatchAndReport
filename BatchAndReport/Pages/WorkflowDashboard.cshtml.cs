using BatchAndReport.Entities;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

public class WorkflowDashboardModel : PageModel
{
    private readonly K2DBContext_Workflow _k2context_workflow;

    public WorkflowDashboardModel(K2DBContext_Workflow context)
    {
        _k2context_workflow = context;
    }

    public List<WorkflowChartDonutDto> DonutProcessChart { get; set; } = new();
    public List<WorkflowChartDonutDto> DonutImproveChart { get; set; } = new();
    public List<WorkflowChartBarDto> BarChartData { get; set; } = new();
    public List<WorkflowGridDto> WorkflowGrid { get; set; } = new();

    public class WorkflowChartDonutDto
    {
        public string Label { get; set; }
        public int Value { get; set; }
        public string Color { get; set; }
    }

    public class WorkflowChartBarDto
    {
        public string Code { get; set; }
        public int PreviousYear { get; set; }
        public int CurrentYear { get; set; }
        public string Description { get; set; }
        public int Year { get; set; }
        public double PercentChange { get; set; }
    }

    public class WorkflowGridDto
    {
        public int No { get; set; }
        public string Code { get; set; }
        public string WorkflowName { get; set; }
        public int PreviousYearDays { get; set; }
        public int CurrentYearDays { get; set; }
        public int DayDifference { get; set; }
        public string PercentDifference { get; set; }
    }

    public async Task OnGetAsync()
    {
        DonutProcessChart = await GetDonutProcessChartFromDb();
        DonutImproveChart = await GetDonutImproveChartFromDb();
        //BarChartData = GetBarChartMockData();
        BarChartData = await GetBarChartDataFromDb();
        //WorkflowGrid = GetWorkflowGridMockData();
        WorkflowGrid = await GetWorkflowGridFromDb();
    }

    private async Task<List<WorkflowChartDonutDto>> GetDonutProcessChartFromDb()
    {
        var result = new List<WorkflowChartDonutDto>();
        var conn = _k2context_workflow.Database.GetDbConnection();
        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_WF_GET_DASHBOARD_WORKPLAN", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@FISCAL_YEAR_ID", DBNull.Value);
        command.Parameters.AddWithValue("@OWNER_BusinessUnitId", DBNull.Value);
        command.Parameters.AddWithValue("@PROCESS_GROUP_CODE", DBNull.Value);

        await connection.OpenAsync();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            result.Add(new WorkflowChartDonutDto
            {
                Label = reader["PLAN_CATEGORIES_NAME"].ToString(),
                Value = Convert.ToInt32(reader["DetailCount"]),
                Color = "#478CC5"
            });
        }
        return result;
    }

    private async Task<List<WorkflowChartDonutDto>> GetDonutImproveChartFromDb()
    {
        var result = new List<WorkflowChartDonutDto>();
        var conn = _k2context_workflow.Database.GetDbConnection();
        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_WF_GET_DASHBOARD_WORKPROCESS", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@FISCAL_YEAR_ID", DBNull.Value);
        command.Parameters.AddWithValue("@OWNER_BusinessUnitId", DBNull.Value);
        command.Parameters.AddWithValue("@PROCESS_GROUP_CODE", DBNull.Value);

        await connection.OpenAsync();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            result.Add(new WorkflowChartDonutDto
            {
                Label = reader["PROCESS_REVIEW_TYPE_NAME"].ToString(),
                Value = Convert.ToInt32(reader["DetailCount"]),
                Color = "#7EB4D7"
            });
        }
        return result;
    }

    private async Task<List<WorkflowChartBarDto>> GetBarChartDataFromDb()
    {
        var result = new List<WorkflowChartBarDto>();
        var conn = _k2context_workflow.Database.GetDbConnection();

        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_WF_GET_DASHBOARD_COMPARESUBPROCESS", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        await connection.OpenAsync();
        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            var code = reader["PROCESS_CODE"]?.ToString();
            var name = reader["PROCESS_NAME"]?.ToString();

            var currentYear = reader["CurrentFiscalYear"]?.ToString() ?? "";
            var previousYear = reader["RefFiscalYear"]?.ToString() ?? "";

            int prevDays = reader["RefProcessDay"] != DBNull.Value ? Convert.ToInt32(reader["RefProcessDay"]) : 0;
            int currDays = reader["CurrentProcessDay"] != DBNull.Value ? Convert.ToInt32(reader["CurrentProcessDay"]) : 0;
            int diffDays = reader["ProcessDayDiff"] != DBNull.Value ? Convert.ToInt32(reader["ProcessDayDiff"]) : 0;
            double percent = reader["ProcessDayDiffPercent"] != DBNull.Value ? Convert.ToDouble(reader["ProcessDayDiffPercent"]) : 0;

            result.Add(new WorkflowChartBarDto
            {
                Code = code,
                Description = name,
                Year = int.TryParse(currentYear, out int y) ? y : 0,
                PreviousYear = prevDays,
                CurrentYear = currDays,
                PercentChange = percent
            });
        }

        return result;
    }

    private async Task<List<WorkflowGridDto>> GetWorkflowGridFromDb()
    {
        var result = new List<WorkflowGridDto>();
        var conn = _k2context_workflow.Database.GetDbConnection();

        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_WF_GET_DASHBOARD_COMPARESUBPROCESS", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        await connection.OpenAsync();
        using var reader = await command.ExecuteReaderAsync();

        int index = 1;

        while (await reader.ReadAsync())
        {
            var code = reader["PROCESS_CODE"]?.ToString();
            var name = reader["PROCESS_NAME"]?.ToString();

            int prev = reader["RefProcessDay"] != DBNull.Value ? Convert.ToInt32(reader["RefProcessDay"]) : 0;
            int curr = reader["CurrentProcessDay"] != DBNull.Value ? Convert.ToInt32(reader["CurrentProcessDay"]) : 0;
            int diff = reader["ProcessDayDiff"] != DBNull.Value ? Convert.ToInt32(reader["ProcessDayDiff"]) : 0;
            double percent = reader["ProcessDayDiffPercent"] != DBNull.Value ? Convert.ToDouble(reader["ProcessDayDiffPercent"]) : 0;

            result.Add(new WorkflowGridDto
            {
                No = index++,
                Code = code,
                WorkflowName = name,
                PreviousYearDays = prev,
                CurrentYearDays = curr,
                DayDifference = diff,
                PercentDifference = percent.ToString("0.##") + "%"
            });
        }

        return result;
    }
    private double CalculatePercentChange(int previous, int current)
    {
        if (previous == 0)
            return 0;
        return Math.Round((double)(previous - current) / previous * 100, 2);
    }


    private List<WorkflowChartBarDto> GetBarChartMockData() => new()
    {
        new() { Code = "P01", PreviousYear = 25, CurrentYear = 20, Description = "กระบวนการ A", Year = 2567, PercentChange = 20 },
        new() { Code = "P02", PreviousYear = 30, CurrentYear = 28, Description = "กระบวนการ B", Year = 2567, PercentChange = 6.67 },
        new() { Code = "P03", PreviousYear = 40, CurrentYear = 35, Description = "กระบวนการ C", Year = 2567, PercentChange = 12.5 },
        new() { Code = "P04", PreviousYear = 50, CurrentYear = 60, Description = "กระบวนการ D", Year = 2567, PercentChange = -20 },
        new() { Code = "P05", PreviousYear = 35, CurrentYear = 30, Description = "กระบวนการ E", Year = 2567, PercentChange = 14.29 },
        new() { Code = "P06", PreviousYear = 45, CurrentYear = 38, Description = "กระบวนการ F", Year = 2567, PercentChange = 15.56 },
        new() { Code = "P07", PreviousYear = 20, CurrentYear = 25, Description = "กระบวนการ G", Year = 2567, PercentChange = -25 },
        new() { Code = "P08", PreviousYear = 60, CurrentYear = 50, Description = "กระบวนการ H", Year = 2567, PercentChange = 16.67 },
        new() { Code = "P09", PreviousYear = 55, CurrentYear = 52, Description = "กระบวนการ I", Year = 2567, PercentChange = 5.45 },
        new() { Code = "P10", PreviousYear = 33, CurrentYear = 29, Description = "กระบวนการ J", Year = 2567, PercentChange = 12.12 }
    };

    private List<WorkflowGridDto> GetWorkflowGridMockData() => new()
    {
        new() { No = 1, Code = "P01", WorkflowName = "กระบวนการ A", PreviousYearDays = 25, CurrentYearDays = 20, DayDifference = -5, PercentDifference = "20%" },
        new() { No = 2, Code = "P02", WorkflowName = "กระบวนการ B", PreviousYearDays = 30, CurrentYearDays = 28, DayDifference = -2, PercentDifference = "6.67%" },
        new() { No = 3, Code = "P03", WorkflowName = "กระบวนการ C", PreviousYearDays = 40, CurrentYearDays = 35, DayDifference = -5, PercentDifference = "12.5%" },
        new() { No = 4, Code = "P04", WorkflowName = "กระบวนการ D", PreviousYearDays = 50, CurrentYearDays = 60, DayDifference = 10, PercentDifference = "-20%" },
        new() { No = 5, Code = "P05", WorkflowName = "กระบวนการ E", PreviousYearDays = 35, CurrentYearDays = 30, DayDifference = -5, PercentDifference = "14.29%" }
    };
}