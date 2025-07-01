using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Collections.Generic;

public class WorkflowDashboardModel : PageModel
{
    public List<WorkflowChartDonutDto> DonutProcessChart { get; set; } = new();
    public List<WorkflowChartDonutDto> DonutImproveChart { get; set; } = new();
    public List<WorkflowChartBarDto> BarChartData { get; set; } = new();
    public List<WorkflowGridDto> WorkflowGrid { get; set; } = new();

    public WorkflowDashboardModel()
    {
        DonutProcessChart = GetDonutProcessChart();
        DonutImproveChart = GetDonutImproveChart();
        BarChartData = GetBarChartData();
        WorkflowGrid = GetWorkflowGrid();
    }

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


    private List<WorkflowChartDonutDto> GetDonutProcessChart() => new()
    {
        new() { Label = "งานบริหาร", Value = 6, Color = "#003F88" },
        new() { Label = "งานส่งเสริมและประสานงานเครือข่าย", Value = 6, Color = "#1963B3" },
        new() { Label = "งานให้บริการ SMEs ครบวงจร", Value = 6, Color = "#478CC5" },
        new() { Label = "งานข้อมูลและสถานการณ์", Value = 6, Color = "#7EB4D7" },
        new() { Label = "งานนโยบายและยุทธศาสตร์", Value = 6, Color = "#B4D6E8" }
    };

    private List<WorkflowChartDonutDto> GetDonutImproveChart() => new()
    {
        new() { Label = "ทบทวนตาม JD", Value = 12, Color = "#003F88" },
        new() { Label = "ทบทวนตาม คจ.2", Value = 12, Color = "#478CC5" },
        new() { Label = "ทบทวนตามการทำงานปัจจุบัน", Value = 36, Color = "#B4D6E8" }
    };

    private List<WorkflowChartBarDto> GetBarChartData() => new()
    {
        new() { Code = "C2.1", PreviousYear = 25, CurrentYear = 23, Description = "การจัดทำแผนปฏิบัติการส่งเสริม SME", Year = 2567, PercentChange = 8 },
        new() { Code = "C2.6", PreviousYear = 42, CurrentYear = 38, Description = "การส่งเสริมการตลาด", Year = 2567, PercentChange = 9.52 },
        new() { Code = "C2.7", PreviousYear = 38, CurrentYear = 30, Description = "พัฒนาเทคโนโลยี SME", Year = 2567, PercentChange = 21.05 },
        new() { Code = "C2.8", PreviousYear = 75, CurrentYear = 60, Description = "การเงินและสินเชื่อ", Year = 2567, PercentChange = 20 },
        new() { Code = "C2.9", PreviousYear = 62, CurrentYear = 58, Description = "ธุรกิจเริ่มต้น", Year = 2567, PercentChange = 6.45 }
    };

    private List<WorkflowGridDto> GetWorkflowGrid() => new()
    {
        new() { No = 11, Code = "C3.2", WorkflowName = "การจัดทำแผนปฏิบัติการส่งเสริม SME", PreviousYearDays = 30, CurrentYearDays = 29, DayDifference = 1, PercentDifference = "3.33%" },
        new() { No = 12, Code = "C3.3", WorkflowName = "การจัดทำแผนปฏิบัติการส่งเสริม SME", PreviousYearDays = 27, CurrentYearDays = 28, DayDifference = -1, PercentDifference = "3.70%" },
        new() { No = 13, Code = "C3.4", WorkflowName = "การจัดทำแผนปฏิบัติการส่งเสริม SME", PreviousYearDays = 25, CurrentYearDays = 27, DayDifference = -2, PercentDifference = "8.00%" }
    };
}
