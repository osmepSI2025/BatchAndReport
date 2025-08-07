using BatchAndReport.Entities;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

public class EContractDashboardModel : PageModel
{
    private readonly K2DBContext_EContract _k2context_econtract;

    public EContractDashboardModel(K2DBContext_EContract context)
    {
        _k2context_econtract = context;
    }

    public List<EContractChartDonutDto> ContractTypeChart { get; set; } = new();
    public List<EContractChartDonutDto> DocumentStatusChart { get; set; } = new();
    public List<EContractChartDonutDto> ContractStatusChart { get; set; } = new();
    public List<LegalKpiDto> LegalKpiChart { get; set; } = new();

    public class EContractChartDonutDto
    {
        public string Label { get; set; }
        public int Value { get; set; }
        public string Color { get; set; }
    }

    public class LegalKpiDto
    {
        public string Owner { get; set; }
        public int Pending { get; set; }
        public int Completed { get; set; }
        public int Total { get; set; }
    }

    public async Task OnGetAsync()
    {
        ContractTypeChart = await GetChartAsync("SP_COUNT_REQUEST_BY_CONTRACT_TYPE", "Contract_Type");
        DocumentStatusChart = await GetChartAsync("SP_COUNT_REQUEST_BY_DOC_STATUS", "Status_Th");
        ContractStatusChart = await GetChartAsync("SP_COUNT_REQUEST_BY_LOOKUP_TYPE", "Contract_Status_Th");
        LegalKpiChart = await GetLegalKpiChartAsync();
    }

    private async Task<List<EContractChartDonutDto>> GetChartAsync(string procedureName, string labelColumn)
    {
        var result = new List<EContractChartDonutDto>();
        var conn = _k2context_econtract.Database.GetDbConnection();

        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand(procedureName, connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        await connection.OpenAsync();
        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            result.Add(new EContractChartDonutDto
            {
                Label = reader[labelColumn].ToString(),
                Value = Convert.ToInt32(reader["TotalRequestCount"]),
                Color = ""
            });
        }

        return result;
    }

    private async Task<List<LegalKpiDto>> GetLegalKpiChartAsync()
    {
        var result = new List<LegalKpiDto>();
        var conn = _k2context_econtract.Database.GetDbConnection();

        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_COUNT_REQUEST_STATUS_BY_OWNER", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        await connection.OpenAsync();
        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            result.Add(new LegalKpiDto
            {
                Owner = reader["OWNER"].ToString(),
                Pending = Convert.ToInt32(reader["รอตรวจสอบ"]),
                Completed = Convert.ToInt32(reader["ตรวจสอบเสร็จสิ้น"]),
                Total = Convert.ToInt32(reader["รวมทั้งสิ้น"])
            });
        }

        return result;
    }
}
