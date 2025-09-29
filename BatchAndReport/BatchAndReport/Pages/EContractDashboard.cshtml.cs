using BatchAndReport.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Threading.Tasks;

public class EContractDashboardModel : PageModel
{
    private readonly K2DBContext_EContract _k2context_econtract;

    public EContractDashboardModel(K2DBContext_EContract context)
    {
        _k2context_econtract = context;
    }

    [BindProperty(SupportsGet = true)]
    public int? Year { get; set; }            // รับ ?Year=2568 หรือ 2025

    [BindProperty(SupportsGet = true)]
    public string? MonthName { get; set; }     // รับ ?MonthName=กันยายน เป็นต้น

    public List<EContractChartDonutDto> ContractTypeChart { get; set; } = new();
    public List<EContractChartDonutDto> DocumentStatusChart { get; set; } = new();
    public List<EContractChartDonutDto> ContractStatusChart { get; set; } = new();
    public List<LegalKpiDto> LegalKpiChart { get; set; } = new();

    // ใช้สำหรับอ่านผลจาก SP_COUNT_REQUEST_BY_CONTRACT_CATEGORY
    public List<ContractCategorySliceDto> ContractCategoryChart { get; set; } = new();

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

    public class ContractCategorySliceDto
    {
        public string Type { get; set; }   // WF_TYPE
        public string Label { get; set; }  // Contract_Category
        public int Value { get; set; }     // TotalRequestCount
    }

    public async Task OnGetAsync()
    {
        // ทุก SP รองรับ @Year (พ.ศ./ค.ศ.) และ @MonthName (ชื่อเดือนภาษาไทย) แล้ว
        ContractTypeChart = await GetChartAsync("SP_COUNT_REQUEST_BY_CONTRACT_TYPE", "Contract_Type", Year, MonthName);
        DocumentStatusChart = await GetChartAsync("SP_COUNT_REQUEST_BY_DOC_STATUS", "Status_Th", Year, MonthName);
        ContractStatusChart = await GetChartAsync("SP_COUNT_REQUEST_BY_LOOKUP_TYPE", "Contract_Status_Th", Year, MonthName);
        LegalKpiChart = await GetLegalKpiChartAsync(Year, MonthName);
        ContractCategoryChart = await GetContractCategoryChartAsync(Year, MonthName);
    }

    private async Task<List<EContractChartDonutDto>> GetChartAsync(string procedureName, string labelColumn, int? year, string? monthName)
    {
        var result = new List<EContractChartDonutDto>();
        var conn = _k2context_econtract.Database.GetDbConnection();

        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand(procedureName, connection) { CommandType = CommandType.StoredProcedure };

        command.Parameters.Add(new SqlParameter("@Year", SqlDbType.Int) { Value = (object?)year ?? DBNull.Value });
        command.Parameters.Add(new SqlParameter("@MonthName", SqlDbType.NVarChar, 30) { Value = (object?)monthName ?? DBNull.Value });

        await connection.OpenAsync();
        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            result.Add(new EContractChartDonutDto
            {
                Label = reader[labelColumn]?.ToString(),
                Value = Convert.ToInt32(reader["TotalRequestCount"]),
                Color = ""
            });
        }
        return result;
    }

    private async Task<List<ContractCategorySliceDto>> GetContractCategoryChartAsync(int? year, string? monthName)
    {
        var result = new List<ContractCategorySliceDto>();
        var conn = _k2context_econtract.Database.GetDbConnection();

        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_COUNT_REQUEST_BY_CONTRACT_CATEGORY", connection)
        { CommandType = CommandType.StoredProcedure };

        command.Parameters.Add(new SqlParameter("@Year", SqlDbType.Int) { Value = (object?)year ?? DBNull.Value });
        command.Parameters.Add(new SqlParameter("@MonthName", SqlDbType.NVarChar, 30) { Value = (object?)monthName ?? DBNull.Value });

        await connection.OpenAsync();
        using var reader = await command.ExecuteReaderAsync();

        // คาดหวังคอลัมน์: WF_TYPE, Contract_Category, TotalRequestCount
        int ordType = reader.GetOrdinal("WF_TYPE");
        int ordLabel = reader.GetOrdinal("Contract_Category");
        int ordVal = reader.GetOrdinal("TotalRequestCount");

        while (await reader.ReadAsync())
        {
            result.Add(new ContractCategorySliceDto
            {
                Type = reader.IsDBNull(ordType) ? "" : reader.GetString(ordType),
                Label = reader.IsDBNull(ordLabel) ? "ไม่ระบุ" : reader.GetString(ordLabel),
                Value = reader.IsDBNull(ordVal) ? 0 : Convert.ToInt32(reader.GetValue(ordVal))
            });
        }
        return result;
    }

    private async Task<List<LegalKpiDto>> GetLegalKpiChartAsync(int? year, string? monthName)
    {
        var result = new List<LegalKpiDto>();
        var conn = _k2context_econtract.Database.GetDbConnection();

        await using var connection = new SqlConnection(conn.ConnectionString);
        await using var command = new SqlCommand("SP_COUNT_REQUEST_STATUS_BY_OWNER", connection)
        { CommandType = CommandType.StoredProcedure };

        command.Parameters.Add(new SqlParameter("@Year", SqlDbType.Int) { Value = (object?)year ?? DBNull.Value });
        command.Parameters.Add(new SqlParameter("@MonthName", SqlDbType.NVarChar, 30) { Value = (object?)monthName ?? DBNull.Value });

        await connection.OpenAsync();
        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            result.Add(new LegalKpiDto
            {
                Owner = reader["OWNER"]?.ToString(),
                Pending = Convert.ToInt32(reader["รอตรวจสอบ"]),
                Completed = Convert.ToInt32(reader["ตรวจสอบเสร็จสิ้น"]),
                Total = Convert.ToInt32(reader["รวมทั้งสิ้น"])
            });
        }
        return result;
    }

    // TXT Export
    public async Task<IActionResult> OnGetExportTxtAsync(int? year, string? monthName)
    {
        if (LegalKpiChart == null || LegalKpiChart.Count == 0)
            LegalKpiChart = await GetLegalKpiChartAsync(year, monthName);

        var sb = new StringBuilder();
        sb.AppendLine("เจ้าหน้าที่\tรอตรวจสอบ\tตรวจสอบเสร็จสิ้น\tรวมทั้งสิ้น");
        foreach (var item in LegalKpiChart)
            sb.AppendLine($"{item.Owner}\t{item.Pending}\t{item.Completed}\t{item.Total}");

        var bytes = Encoding.UTF8.GetBytes(sb.ToString());
        var suffix = BuildSuffix(year, monthName);
        return File(bytes, "text/plain", $"LegalKPI{suffix}.txt");
    }

    // XLS Export
    public async Task<IActionResult> OnGetExportXlsAsync(int? year, string? monthName)
    {
        if (LegalKpiChart == null || LegalKpiChart.Count == 0)
            LegalKpiChart = await GetLegalKpiChartAsync(year, monthName);

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("Legal KPI");

        // Header
        worksheet.Cells[1, 1].Value = "เจ้าหน้าที่";
        worksheet.Cells[1, 2].Value = "รอตรวจสอบ";
        worksheet.Cells[1, 3].Value = "ตรวจสอบเสร็จสิ้น";
        worksheet.Cells[1, 4].Value = "รวมทั้งสิ้น";

        // Data
        int row = 2;
        foreach (var item in LegalKpiChart)
        {
            worksheet.Cells[row, 1].Value = item.Owner;
            worksheet.Cells[row, 2].Value = item.Pending;
            worksheet.Cells[row, 3].Value = item.Completed;
            worksheet.Cells[row, 4].Value = item.Total;
            row++;
        }

        var bytes = package.GetAsByteArray();
        var suffix = BuildSuffix(year, monthName);
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"LegalKPI{suffix}.xlsx");
    }

    // JPEG Export (stub)
    public async Task<IActionResult> OnGetExportJpegAsync()
    {
        var bytes = Encoding.UTF8.GetBytes("JPEG export not implemented yet.");
        return File(bytes, "image/jpeg", "LegalKPI.jpg");
    }

    private static string BuildSuffix(int? year, string? monthName)
    {
        var parts = new List<string>();
        if (year != null) parts.Add(year!.Value.ToString());
        if (!string.IsNullOrWhiteSpace(monthName)) parts.Add(monthName!.Trim());
        return parts.Count > 0 ? "_" + string.Join("_", parts) : "";
    }
}
