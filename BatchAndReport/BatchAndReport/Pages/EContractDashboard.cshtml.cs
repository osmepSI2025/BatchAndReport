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
    // TXT Export
    public async Task<IActionResult> OnGetExportTxtAsync(string type)
    {
        if (LegalKpiChart == null || LegalKpiChart.Count == 0)
            LegalKpiChart = await GetLegalKpiChartAsync();

        var sb = new StringBuilder();
        sb.AppendLine("เจ้าหน้าที่\tรอตรวจสอบ\tตรวจสอบเสร็จสิ้น\tรวมทั้งสิ้น");
        foreach (var item in LegalKpiChart)
            sb.AppendLine($"{item.Owner}\t{item.Pending}\t{item.Completed}\t{item.Total}");

        var bytes = Encoding.UTF8.GetBytes(sb.ToString());
        return File(bytes, "text/plain", "LegalKPI.txt");
    }


    // XLS Export (stub)
    public async Task<IActionResult> OnGetExportXlsAsync()
    {
        if (LegalKpiChart == null || LegalKpiChart.Count == 0)
            LegalKpiChart = await GetLegalKpiChartAsync();

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set license context

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
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "LegalKPI.xlsx");
    }

    // PDF Export (stub)
    //public async Task<IActionResult> OnGetExportPdfAsync()
    //{
    //    // TODO: Implement PDF export logic
    //    // For now, return a placeholder file
    //    var bytes = Encoding.UTF8.GetBytes("PDF export not implemented yet.");
    //    return File(bytes, "application/pdf", "LegalKPI.pdf");
    //}

    // JPEG Export (stub)
    public async Task<IActionResult> OnGetExportJpegAsync()
    {
        // TODO: Implement JPEG export logic
        // For now, return a placeholder file
        var bytes = Encoding.UTF8.GetBytes("JPEG export not implemented yet.");
        return File(bytes, "image/jpeg", "LegalKPI.jpg");
    }
}
