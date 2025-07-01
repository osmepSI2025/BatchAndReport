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

    public List<BudgetData> BudgetChart { get; set; } = new();
    public List<SupportIssue> SupportChart { get; set; } = new();
    public List<ProjectType> TypeChart { get; set; } = new();
    public List<RegionData> RegionChart { get; set; } = new();

    public class BudgetData
    {
        public string Year { get; set; } = "";
        public int RequestBudget { get; set; }
        public int ApproveBudget { get; set; }
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

    private async Task<List<BudgetData>> GetBudgetChartAsync()
    {
        var result = new List<BudgetData>();

        var conn = _k2context_sme.Database.GetDbConnection();
        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_SME_GET_BUDGET_CHART", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        await connection.OpenAsync();

        using var reader = await command.ExecuteReaderAsync();
        var yearColumns = new List<string>();

        for (int i = 0; i < reader.FieldCount; i++)
        {
            var colName = reader.GetName(i);
            if (colName != "column type")
                yearColumns.Add(colName);
        }

        var tempDict = new Dictionary<string, BudgetData>();

        while (await reader.ReadAsync())
        {
            var budgetType = reader["column type"].ToString();

            foreach (var year in yearColumns)
            {
                if (!int.TryParse(reader[year]?.ToString(), out int value))
                    continue;

                if (!tempDict.ContainsKey(year))
                    tempDict[year] = new BudgetData { Year = year };

                if (budgetType == "pass_budget")
                    tempDict[year].ApproveBudget = value;
                else
                    tempDict[year].RequestBudget = value;
            }
        }

        result = tempDict.Values.ToList();
        return result;
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
                Count = Convert.ToInt32(reader["department_count"])
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
}
