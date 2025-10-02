using BatchAndReport.Entities;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

public class SMEDashboardModel : PageModel
{
    private readonly K2DBContext_SME _k2context_sme;

    public SMEDashboardModel(K2DBContext_SME context)
    {
        _k2context_sme = context;
    }

    public List<BudgetRow> BudgetChart { get; set; } = new();
    public List<SupportIssue> SupportChart { get; set; } = new();
    public List<ProjectType> TypeChart { get; set; } = new();
    public List<RegionData> RegionChart { get; set; } = new();

    public class BudgetRow
    {
        public string Year { get; set; } = "TOTAL";  // ปรับเป็นปีจริงได้ถ้ามี
        public decimal RequestBudget { get; set; }   // budget_req (ล.บ.)
        public decimal ApproveBudget { get; set; }   // budget_req_pass (ล.บ.)
    }
    private static object DbNullIf<T>(T? v) where T : struct
    => v.HasValue ? (object)v.Value : DBNull.Value;

    private static object DbNullIf(string? v)
        => string.IsNullOrWhiteSpace(v) ? DBNull.Value : v!;

    private static decimal SafeDecimal(SqlDataReader reader, string col)
    {
        if (reader[col] == DBNull.Value) return 0m;
        try { return Convert.ToDecimal(reader[col]); } catch { return 0m; }
    }

    public class SupportIssue { public string Issue { get; set; } = ""; public int Count { get; set; } }
    public class ProjectType { public string Type { get; set; } = ""; public int Percent { get; set; } }
    public class RegionData { public string Region { get; set; } = ""; public int Count { get; set; } }

    public async Task OnGetAsync()
    {
        BudgetChart = await GetBudgetChartAsync();
        SupportChart = await GetSupportChartAsync();
        TypeChart = await GetTypeChartAsync();
        RegionChart = await GetRegionChartAsync();
    }

    private async Task<List<BudgetRow>> GetBudgetChartAsync(
    int? fiscalYearId = null,
    string? departmentCode = null,
    int? operationAreaId = null,
    string? budgetSourceCode = null)
    {
        var rows = new List<BudgetRow>();

        var conn = _k2context_sme.Database.GetDbConnection();
        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_SME_GET_DASHBOARD_DETAIL", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@FISCAL_YEAR_ID", DbNullIf(fiscalYearId));
        command.Parameters.AddWithValue("@DEPARTMENT_CODE", DbNullIf(departmentCode));
        command.Parameters.AddWithValue("@OPERATION_AREA_ID", DbNullIf(operationAreaId));
        command.Parameters.AddWithValue("@BUDGET_SOURCE_CODE", DbNullIf(budgetSourceCode));

        await connection.OpenAsync();
        using var reader = await command.ExecuteReaderAsync();

        if (await reader.ReadAsync())
        {
            rows.Add(new BudgetRow
            {
                Year = "TOTAL",
                RequestBudget = SafeDecimal(reader, "budget_req"),
                ApproveBudget = SafeDecimal(reader, "budget_req_pass")
            });
        }

        return rows;
    }

    private async Task<List<SupportIssue>> GetSupportChartAsync()
    {
        var result = new List<SupportIssue>();
        var conn = _k2context_sme.Database.GetDbConnection();
        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_SME_GET_STRATEGY_CHART", connection)
        {
            CommandType = CommandType.StoredProcedure
        };
        await connection.OpenAsync();

        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            result.Add(new SupportIssue
            {
                Issue = reader["TOPIC"].ToString(),
                Count = SafeToInt(reader["department_count"])
            });
        }
        return result;
    }

    private async Task<List<ProjectType>> GetTypeChartAsync()
    {
        var result = new List<ProjectType>();
        var conn = _k2context_sme.Database.GetDbConnection();
        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_SME_GET_PROJECT_STATUS_CHART", connection)
        {
            CommandType = CommandType.StoredProcedure
        };
        await connection.OpenAsync();

        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            result.Add(new ProjectType
            {
                Type = reader["status"].ToString(),
                Percent = Convert.ToInt32(reader["project_status_count"])
            });
        }
        return result;
    }

    private async Task<List<RegionData>> GetRegionChartAsync()
    {
        var result = new List<RegionData>();
        var conn = _k2context_sme.Database.GetDbConnection();
        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_SME_GET_AREA_CHART", connection)
        {
            CommandType = CommandType.StoredProcedure
        };
        await connection.OpenAsync();

        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            result.Add(new RegionData
            {
                Region = reader["OPERATION_AREA_NAME"].ToString(),
                Count = Convert.ToInt32(reader["sum_project"])
            });
        }
        return result;
    }

    private int SafeToInt(object obj)
    {
        return obj != DBNull.Value ? Convert.ToInt32(obj) : 0;
    }
}
