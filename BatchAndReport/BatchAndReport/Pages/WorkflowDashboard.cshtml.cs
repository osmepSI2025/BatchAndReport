using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using BatchAndReport.Entities;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;

public class WorkflowDashboardModel : PageModel
{
    private readonly K2DBContext_Workflow _k2context_workflow;

    public WorkflowDashboardModel(K2DBContext_Workflow context)
    {
        _k2context_workflow = context;
    }

    // ===== Filter Parameters =====
    public int? FiscalYearId { get; set; }
    public string? OwnerBusinessUnitId { get; set; }
    public string? ProcessTypeCode { get; set; }

    // ===== Chart / Grid Data =====
    public List<WorkflowChartDonutDto> DonutProcessChart { get; set; } = new();
    public List<WorkflowChartDonutDto> DonutImproveChart { get; set; } = new();
    public List<WorkflowChartBarDto> BarChartData { get; set; } = new();
    public List<WorkflowGridDto> WorkflowGrid { get; set; } = new();

    // ===== DTOs =====
    public class WorkflowChartDonutDto
    {
        public string Label { get; set; } = "";
        public int Value { get; set; }
        public string Color { get; set; } = "#478CC5";
    }

    public class WorkflowChartBarDto
    {
        public string Code { get; set; } = "";
        public int PreviousYear { get; set; }
        public int CurrentYear { get; set; }
        public string Description { get; set; } = "";
        public int Year { get; set; }
        public double PercentChange { get; set; }
    }

    public class WorkflowGridDto
    {
        public int No { get; set; }
        public string Code { get; set; } = "";
        public string WorkflowName { get; set; } = "";
        public int PreviousYearDays { get; set; }
        public int CurrentYearDays { get; set; }
        public int DayDifference { get; set; }
        public string PercentDifference { get; set; } = "-";
    }

    public async Task OnGetAsync(int? FiscalYearId, string? OwnerBusinessUnitId, string? ProcessTypeCode)
    {
        // bind query params
        this.FiscalYearId = FiscalYearId;
        this.OwnerBusinessUnitId = string.IsNullOrWhiteSpace(OwnerBusinessUnitId) ? null : OwnerBusinessUnitId;
        this.ProcessTypeCode = string.IsNullOrWhiteSpace(ProcessTypeCode) ? null : ProcessTypeCode;

        try { DonutProcessChart = await GetDonutProcessChartFromDb(); } catch { DonutProcessChart = new(); }
        try { DonutImproveChart = await GetDonutImproveChartFromDb(); } catch { DonutImproveChart = new(); }
        try { BarChartData = await GetBarChartDataFromDb(); } catch { BarChartData = new(); }
        try { WorkflowGrid = await GetWorkflowGridFromDb(); } catch { WorkflowGrid = new(); }
    }

    private object ToDbNullable<T>(T? value) where T : struct
        => value.HasValue ? (object)value.Value : DBNull.Value;

    private object ToDbNullable(string? value)
        => string.IsNullOrWhiteSpace(value) ? DBNull.Value : value!;

    // ====== Donut: Workplan ======
    private async Task<List<WorkflowChartDonutDto>> GetDonutProcessChartFromDb()
    {
        var result = new List<WorkflowChartDonutDto>();
        var conn = _k2context_workflow.Database.GetDbConnection();

        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_WF_GET_DASHBOARD_WORKPLAN", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@FISCAL_YEAR_ID", ToDbNullable(FiscalYearId));
        command.Parameters.AddWithValue("@OWNER_BusinessUnitId", ToDbNullable(OwnerBusinessUnitId));
        command.Parameters.AddWithValue("@PROCESS_TYPE_CODE", ToDbNullable(ProcessTypeCode));

        await connection.OpenAsync();
        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            var label = SafeString(reader, "PLAN_CATEGORIES_NAME");
            var value = SafeInt(reader, "DetailCount");
            if (!string.IsNullOrWhiteSpace(label))
            {
                result.Add(new WorkflowChartDonutDto
                {
                    Label = label,
                    Value = value,
                    Color = "#478CC5"
                });
            }
        }

        return result;
    }

    // ====== Donut: Workprocess (Review Type) ======
    private async Task<List<WorkflowChartDonutDto>> GetDonutImproveChartFromDb()
    {
        var result = new List<WorkflowChartDonutDto>();
        var conn = _k2context_workflow.Database.GetDbConnection();

        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_WF_GET_DASHBOARD_WORKPROCESS", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@FISCAL_YEAR_ID", ToDbNullable(FiscalYearId));
        command.Parameters.AddWithValue("@OWNER_BusinessUnitId", ToDbNullable(OwnerBusinessUnitId));
        command.Parameters.AddWithValue("@PROCESS_TYPE_CODE", ToDbNullable(ProcessTypeCode));

        await connection.OpenAsync();
        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            var label = SafeString(reader, "PROCESS_REVIEW_TYPE_NAME");
            var value = SafeInt(reader, "DetailCount");
            if (!string.IsNullOrWhiteSpace(label))
            {
                result.Add(new WorkflowChartDonutDto
                {
                    Label = label,
                    Value = value,
                    Color = "#7EB4D7"
                });
            }
        }

        return result;
    }

    // ====== Bar: Top 10 ======
    private async Task<List<WorkflowChartBarDto>> GetBarChartDataFromDb()
    {
        var rows = new List<WorkflowChartBarDto>();
        var conn = _k2context_workflow.Database.GetDbConnection();

        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_WF_GET_DASHBOARD_COMPARESUBPROCESS", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@FISCAL_YEAR_ID", ToDbNullable(FiscalYearId));
        command.Parameters.AddWithValue("@OWNER_BusinessUnitId", ToDbNullable(OwnerBusinessUnitId));
        command.Parameters.AddWithValue("@PROCESS_TYPE_CODE", ToDbNullable(ProcessTypeCode));

        await connection.OpenAsync();
        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            var code = SafeString(reader, "PROCESS_CODE");
            var name = SafeString(reader, "PROCESS_NAME");

            int prevDays = SafeInt(reader, "RefProcessDay");
            int currDays = SafeInt(reader, "CurrentProcessDay");

            double percent = 0;
            if (prevDays > 0)
                percent = Math.Round(((double)currDays - prevDays) / prevDays * 100.0, 2);

            rows.Add(new WorkflowChartBarDto
            {
                Code = code,
                Description = name,
                PreviousYear = prevDays,
                CurrentYear = currDays,
                PercentChange = percent
            });
        }

        // Top 10 by absolute difference
        return rows
            .OrderByDescending(x => Math.Abs(x.CurrentYear - x.PreviousYear))
            .ThenBy(x => x.Description ?? string.Empty)
            .Take(10)
            .ToList();
    }

    // ====== Grid ======
    private async Task<List<WorkflowGridDto>> GetWorkflowGridFromDb()
    {
        var list = new List<WorkflowGridDto>();
        var conn = _k2context_workflow.Database.GetDbConnection();

        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_WF_GET_DASHBOARD_COMPARESUBPROCESS", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@FISCAL_YEAR_ID", ToDbNullable(FiscalYearId));
        command.Parameters.AddWithValue("@OWNER_BusinessUnitId", ToDbNullable(OwnerBusinessUnitId));
        command.Parameters.AddWithValue("@PROCESS_TYPE_CODE", ToDbNullable(ProcessTypeCode));

        await connection.OpenAsync();
        using var reader = await command.ExecuteReaderAsync();

        int index = 1;
        while (await reader.ReadAsync())
        {
            var code = SafeString(reader, "PROCESS_CODE");
            var name = SafeString(reader, "PROCESS_NAME");

            int prev = SafeInt(reader, "RefProcessDay");
            int curr = SafeInt(reader, "CurrentProcessDay");
            int diff = curr - prev;

            string percentText = prev == 0
                ? "-"
                : (Math.Round(((double)curr - prev) / prev * 100.0, 2).ToString("0.##") + "%");

            list.Add(new WorkflowGridDto
            {
                No = index++,
                Code = code,
                WorkflowName = name,
                PreviousYearDays = prev,
                CurrentYearDays = curr,
                DayDifference = diff,
                PercentDifference = percentText
            });
        }

        return list;
    }

    // ===== Helpers =====
    private static string SafeString(SqlDataReader reader, string col)
        => reader[col] == DBNull.Value ? "" : reader[col]?.ToString() ?? "";

    private static int SafeInt(SqlDataReader reader, string col)
    {
        if (reader[col] == DBNull.Value) return 0;
        if (int.TryParse(reader[col]?.ToString(), out var v)) return v;
        try { return Convert.ToInt32(reader[col]); } catch { return 0; }
    }
}
