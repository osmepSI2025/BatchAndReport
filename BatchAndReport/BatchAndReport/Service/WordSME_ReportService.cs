using BatchAndReport.DAO;
using BatchAndReport.Models;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using System.Text;
public class WordSME_ReportService
{
    private readonly WordServiceSetting _w;
    private readonly Econtract_Report_GADAO _e;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    public WordSME_ReportService(WordServiceSetting ws, Econtract_Report_GADAO e
        , IConverter pdfConverter
        )
    {
        _w = ws;
        _e = e;
        _pdfConverter = pdfConverter; // กำหนดค่า DI สำหรับ PDF Converter
    }


    public async Task<byte[]> GenerateSummarySME_Budget_ToPdf(
    List<SMESummaryProjectModels> projects,
    List<SMEStrategyDetailModels> strategyList,
    string year)
    {
        // Read logo and convert to Base64 (if needed in HTML)
        var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
        string logoBase64 = "";
        if (System.IO.File.Exists(logoPath))
        {
            var bytes = await System.IO.File.ReadAllBytesAsync(logoPath);
            logoBase64 = Convert.ToBase64String(bytes);
        }

        // Absolute font path for PDF rendering
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");

        // Build HTML content
        var htmlBody = new StringBuilder();

        // Build HTML content
        var culture = new CultureInfo("th-TH");
        int totalProjects = projects.Sum(p => p.ProjectCount ?? 0);
        decimal totalBudget = projects.Sum(p => p.Budget ?? 0);


        // ------------------ Part 1: Summary ------------------
        htmlBody.Append($@"
    <div class='t-16 text-center'><b>ภาพรวมโครงการและงบประมาณเพื่อการส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (SME)</br>
    ภายใต้แผนปฏิบัติการส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ประจำปี พ.ศ. {year}</b></div>
    <hr>
    <table style='width: 100%;'>
        <tr style='font-weight:bold;background:#f0f0f0;'>
            <th>ประเด็นการส่งเสริม SME</th>
            <th>จำนวนโครงการ</th>
            <th>งบประมาณ (ล้านบาท)</th>
        </tr>
");

        int i = 1;
        foreach (var row in projects)
        {
            htmlBody.Append($@"
        <tr>
            <td>ประเด็นที่{i} {row.IssueName ?? ""}</td>
            <td>{row.ProjectCount?.ToString("N0", culture) ?? "0"}</td>
            <td class='text-right' >{(row.Budget.GetValueOrDefault() / 1_000_000).ToString("N4", culture)}</td>
        </tr>
    ");
            i++;
        }

        htmlBody.Append($@"
    <tr style='font-weight:bold;background:#f0f0f0;'>
        <td>รวมทั้งหมด</td>
        <td>{totalProjects.ToString("N0", culture)}</td>
        <td class='text-right'>{(totalBudget / 1_000_000).ToString("N4", culture)}</td>
    </tr>
</table>
<br>
<div>โดยมีหน่วยงานทั้งหมด {strategyList.Where(p => !string.IsNullOrEmpty(p.Ministry_Id)).Select(p => p.Ministry_Id).Distinct().Count()} กระทรวง {strategyList.Where(p => !string.IsNullOrEmpty(p.DepartmentCode)).Select(p => p.DepartmentCode).Distinct().Count()} หน่วยงาน</div>
<br>
");
        // ------------------ Part 2: Strategy Detail ------------------
        var grouped = strategyList.GroupBy(p => p.Topic).ToList();
        int topicIndex = 1;
        foreach (var topicGroup in grouped)
        {
            htmlBody.Append($@"<div class='t-16'><b>ประเด็นการส่งเสริมที่ {topicIndex} {topicGroup.Key}</b></div>");
            var strategyGrouped = topicGroup.GroupBy(p => p.StrategyDesc).ToList();
            int strategyIndex = 1;
            foreach (var strategyGroup in strategyGrouped)
            {
                var totalProject = strategyGroup.Count();
                var sumBudget = strategyGroup.Sum(p => p.BudgetAmount);

                htmlBody.Append($@"<div >จำนวน {totalProject} โครงการ งบประมาณ {sumBudget:N2} ล้านบาท</div>");

                htmlBody.Append($@"<div class='t-16'><b>กลยุทธ์ที่ {strategyIndex} {strategyGroup.Key}</b></div>");

                // Table
                htmlBody.Append($@"
            <table style='width: 100%;'>
                <tr>
                    <th>หน่วยงาน/โครงการ</th>
                    <th>งบประมาณ</th>
                </tr>
        ");

                var deptGrouped = strategyGroup.GroupBy(p => new { p.DepartmentCode, p.Department }).ToList();
                int projectIndex = 1;
                foreach (var deptGroup in deptGrouped)
                {
                    var deptTotal = deptGroup.Sum(p => p.BudgetAmount);
                    htmlBody.Append($@"
                <tr style='font-weight:bold;background:#e0e0e0;'>
                    <td>{deptGroup.Key.Department}</td>
                    <td class='text-right' >{deptTotal:N2}</td>
                </tr>
            ");
                    foreach (var proj in deptGroup)
                    {
                        htmlBody.Append($@"
                    <tr>
                        <td style='padding-left:32px;'>{projectIndex}. {proj.ProjectName}</td>
                        <td class='text-right'>{proj.BudgetAmount:N2}</td>
                    </tr>
                ");
                        projectIndex++;
                    }
                }

                htmlBody.Append("</table><br>");
                strategyIndex++;
            }
            topicIndex++;
            htmlBody.Append("<div style='page-break-after:always;'></div>");
        }
        // Compose full HTML
        var html = $@"
    <html>
    <head>
        <meta charset='utf-8'>
        <style>
            @font-face {{
                font-family: 'THSarabunNew';
                src: url('file:///{fontPath}') format('truetype');
                font-weight: normal;
                font-style: normal;
            }}
            body {{
                font-size: 22px;
                font-family: 'THSarabunNew', Arial, sans-serif;
            }}
            .t-18 {{ font-size: 1.5em; }}
            .t-18 {{ font-size: 1.7em; }}
            .t-22 {{ font-size: 1.9em; }}
                 .tab1 {{ text-indent: 48px;  word-break: break-all;  }}
        .tab2 {{ text-indent: 96px;  word-break: break-all; }}
        .tab3 {{ text-indent: 144px;  word-break: break-all; }}
        .tab4 {{ text-indent: 192px;  word-break: break-all;}}
            .d-flex {{ display: flex; }}
            .w-100 {{ width: 100%; }}
            .w-40 {{ width: 40%; }}
            .w-50 {{ width: 50%; }}
            .w-60 {{ width: 60%; }}
            .text-center {{ text-align: center; }}
            .sign-single-right {{
                display: flex;
                flex-direction: column;
                position: relative;
                left: 20%;
            }}
            .sign-double {{ display: flex; }}
            .text-center-right-brake {{
                margin-left: 50%;
                word-break: break-all;
            }}
            .text-right {{ text-align: right; }}
            .contract, .section {{
                margin: 12px 0;
                line-height: 1.7;
            }}
            .section {{
                font-weight: bold;
                font-size: 1.2em;
                text-align: left;
                margin-top: 24px;
            }}
            .signature-table {{
                width: 100%;
                margin-top: 32px;
                border-collapse: collapse;
            }}
            .signature-table td {{
                padding: 16px;
                text-align: center;
                vertical-align: top;
                font-size: 1.1em;
            }}
   .table {{ width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 28pt; }}
        .table th, .table td {{ border: 1px solid #000; padding: 8px; }}

        </style>
    </head>
    <body>
        {htmlBody}
    </body>
    </html>
    ";

        var doc = new HtmlToPdfDocument()
        {
            GlobalSettings = {
            PaperSize = PaperKind.A4,
            Orientation = Orientation.Portrait,
            Margins = new MarginSettings
            {
                Top = 20,
                Bottom = 20,
                Left = 20,
                Right = 20
            }
        },
            Objects = {
            new ObjectSettings() {
                HtmlContent = html,
                FooterSettings = new FooterSettings
                {
                    FontName = "THSarabunNew",
                    FontSize = 6,
                    Line = false,
                    Center = "[page] / [toPage]"
                }
            }
        }
        };

        var pdfBytes = _pdfConverter.Convert(doc);
        return pdfBytes;
    }
    public async Task<byte[]> ExportSMEProjectDetail_ToPDF(
SMEProjectDetailModels model
)
    {
        // Read logo and convert to Base64 (if needed in HTML)
        var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
        string logoBase64 = "";
        if (System.IO.File.Exists(logoPath))
        {
            var bytes = await System.IO.File.ReadAllBytesAsync(logoPath);
            logoBase64 = Convert.ToBase64String(bytes);
        }

        // Absolute font path for PDF rendering
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");

        // Build HTML content
        var htmlBody = new StringBuilder();


        // Heading
        htmlBody.Append($@"
        <div class='t-22 text-center'><b>แบบฟอร์มการจัดทำข้อเสนอ ประจำปีงบประมาณ {model.FiscalYearDesc}</b></div>
        <br>
        <div>กระทรวง : {model.MinistryName}</div>
        <div>หน่วยงาน : {model.DepartmentName}</div>
        <div>ชื่อกิจกรรม : {model.ActivityName}</div>
        <div>งบประมาณ : {model.BudgetAmount:N0}</div>
        <br>
        <div><b>□ ใช้งบประมาณ</b></div>
        <div class='tab1'>□ xxxxxxxxxxx</div>
        <div class='tab1'>□ xxxxxxxxxxx</div>
        <div>□ ไม่ใช้งบประมาณ</div>
        <div><b>สถานภาพโครงการ : □ โครงการใหม่ □ โครงการต่อเนื่อง □ โครงการเดิม □ โครงการ Flagship</b></div>
        <br>
    ");

        // Responsible Table
        htmlBody.Append($@"
<table style='width:100%; border-collapse:collapse; border:1px solid #000;'>
    <tr style='background:#e0e0e0; font-weight:bold;'>
                <td style='border:1px solid #000;'></td>
                <td style='border:1px solid #000;text-align:center;'>ผู้รับผิดชอบโครงการ</td>
                <td style='border:1px solid #000;text-align:center;'>ผู้ประสานงาน</td>
            </tr>
            <tr>
                <td style='border:1px solid #000;'>ชื่อ-นามสกุล</td>
                <td style='border:1px solid #000;'>{model.OwnerName}</td>
                <td style='border:1px solid #000;'>{model.ContactName}</td>
            </tr>
            <tr>
                <td style='border:1px solid #000;'>ตำแหน่ง</td>
                <td style='border:1px solid #000;'>{model.OwnerPosition}</td>
                <td style='border:1px solid #000;'>{model.ContactPosition}</td>
            </tr>
            <tr>
                <td style='border:1px solid #000;'>โทรศัพท์</td>
                <td style='border:1px solid #000;'>{model.OwnerPhone}</td>
                <td style='border:1px solid #000;'>{model.ContactPhone}</td>
            </tr>
            <tr>
                <td style='border:1px solid #000;'>มือถือ</td>
                <td style='border:1px solid #000;'>{model.OwnerMobile}</td>
                <td style='border:1px solid #000;'>{model.ContactMobile}</td>
            </tr>
            <tr>
                <td style='border:1px solid #000;'>Email</td>
                <td style='border:1px solid #000;'>{model.OwnerEmail}</td>
                <td style='border:1px solid #000;'>{model.ContactEmail}</td>
            </tr>
            <tr>
                <td style='border:1px solid #000;'>Line ID</td>
                <td style='border:1px solid #000;'>{model.OwnerLineId}</td>
                <td style='border:1px solid #000;'>{model.ContactLineId}</td>
            </tr>
        </table>
        <br>
    ");

        htmlBody.Append($@"
        <div  class='tab1'><b>ประเด็นสำคัญในการส่งเสริม SME ปี พ.ศ.{model.FiscalYearDesc}</b></div>
        <div  class='tab1'><b>□ Digital □ Environment/Green □ Social □ Governance □ Soft power</b></div>
        <div  class='tab1'>ประเด็นสำคัญในการส่งเสริม SME ปี พ.ศ.{model.FiscalYearDesc} ประเด็นการส่งเสริม/กลยุทธ์ที่สอดคล้องกับแผนปฎิบัติการส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อมประจำปีงบประมาณ (เลือกเพียง 1 ประเด็นการส่งเสริม 1 กลยุทธ์ ต่อ 1 โครงการ)</div>
    ");

        // Strategy Table
        htmlBody.Append($@"
        <table style='width:100%; border-collapse:collapse;'>
            <tr style='background:#e0e0e0; font-weight:bold;'>
                <td style='border:1px solid #000;text-align:center;'>ประเด็นการส่งเสริม</td>
                <td style='border:1px solid #000;text-align:center;'>กลยุทธ์</td>
            </tr>
    ");
        if (model.Strategies != null)
        {
            int index = 1;
            foreach (var strategy in model.Strategies)
            {
                htmlBody.Append($@"
                <tr>
                    <td style='border:1px solid #000;'>□ {strategy.StrategyId}</td>
                   <td style='border:1px solid #000;'>□ {index++} {strategy.StrategyDesc}</td>
                </tr>
            ");
            }
        }
        htmlBody.Append("</table><br>");

        htmlBody.Append($@"
        <div  class='tab1'><b>ความสำคัญของโครงการ/หลักการและเหตุผล :</b></div>
        <div  class='tab1'>{model.ProjectRationale ?? ""}</div>
        <div  class='tab1'><b>วัตถุประสงค์ของโครงการ :</b></div>
        <div  class='tab1'>{model.ProjectObjective ?? ""}</div>
        <div  class='tab1'><b>กลุ่มเป้าหมาย (สามารถเลือกได้มากกว่า 1 กลุ่มเป้าหมาย):</b></div>
        <div class='tab2'>□ วิสาหกิจระยะเริ่มต้น Early-Stage Enterprise □ วิสาหกิจขนาดย่อม Small Enterprise</div>
        <div class='tab2'>□ วิสาหกิจรายย่อย Micro Enterprise □ วิสาหกิจขนาดกลาง Medium Enterprise □ ทุกกลุ่ม</div>
        <div  class='tab1'><b>รายละเอียดแผนการดำเนินงาน/กิจกรรม...</b></div>
        <div  class='tab1'>{model.Activities ?? ""}</div>
        <div  class='tab1'><b>จุดเด่นของโครงการ :</b></div>
        <div  class='tab1'>{model.ProjectFocus ?? ""}</div>
        <div  class='tab1'><b>พื้นที่ดำเนินการ :</b></div>
        <div  class='tab1'>{string.Join(", ", model.OperationArea ?? new List<string>())}</div>
        <div  class='tab1'><b>สาขาเป้าหมาย :</b></div>
        <div  class='tab1'>{string.Join(", ", model.IndustrySector ?? new List<string>())}</div>
        <div  class='tab1'><b>การพัฒนา 11 อุตสาหกรรม Soft Power :</b></div>
        <div class='tab2'>□ power 1</div>
        <div class='tab2'>□ power 2</div>
        <div class='tab2'>□ power 3</div>
        <div  class='tab1'><b>ระยะเวลาในการดำเนินโครงการ :</b></div>
        <div  class='tab1'>{model.Timeline ?? ""}</div>
        <div  class='tab1'><b>หน่วยงานที่ร่วมบูรณาการ...</b></div>
        <div  class='tab1'>{model.OrgPartner} ทำหน้าที่ {model.RoleDescription}</div>
        <div  class='tab1'><b>ตัวชี้วัดที่สำคัญ...</b></div>
    ");

        // Output Indicators Table
        htmlBody.Append($@"
   <table style='width:100%; border-collapse:collapse;'>
            <tr style='background:#e0e0e0; font-weight:bold;'>
               <td style='border:1px solid #000;text-align:center;'>ตัวชี้วัดผลผลิต</td>
               <td style='border:1px solid #000;text-align:center;'>จำนวนเป้าหมาย</td>
               <td style='border:1px solid #000;text-align:center;'>หน่วยนับ</td>
               <td style='border:1px solid #000;text-align:center;'>วิธีการวัดผล</td>
            </tr>
    ");
        if (model.OutputIndicators != null)
        {
            foreach (var item in model.OutputIndicators)
            {
                htmlBody.Append($@"
                <tr>
                   <td style='border:1px solid #000;'>{item.Name}</td>
                   <td style='border:1px solid #000;'>{item.Target}</td>
                   <td style='border:1px solid #000;'>{item.Unit}</td>
                   <td style='border:1px solid #000;'>{item.Method}</td>
                </tr>
            ");
            }
        }
        htmlBody.Append("</table><br>");

        // Outcome Indicators Table
        htmlBody.Append($@"
        <table style='width:100%; border-collapse:collapse;'>
            <tr style='background:#e0e0e0; font-weight:bold;'>
               <td style='border:1px solid #000;text-align:center;'>ตัวชี้วัดผลลัพธ์</td>
               <td style='border:1px solid #000;text-align:center;'>จำนวนเป้าหมาย</td>
               <td style='border:1px solid #000;text-align:center;'>หน่วยนับ</td>
               <td style='border:1px solid #000;text-align:center;'>วิธีการวัดผล</td>
            </tr>
    ");
        if (model.OutcomeIndicators != null)
        {
            foreach (var item in model.OutcomeIndicators)
            {
                htmlBody.Append($@"
                <tr>
                   <td style='border:1px solid #000;'>{item.Name}</td>
                   <td style='border:1px solid #000;'>{item.Target}</td>
                   <td style='border:1px solid #000;'>{item.Unit}</td>
                   <td style='border:1px solid #000;'>{item.Method}</td>
                </tr>
            ");
            }
        }
        htmlBody.Append("</table><br>");

        htmlBody.Append($@"
        <div>ข้อมูลอื่นๆ เพิ่มเติม...</div>
        <div>{model.AdditionalNotes ?? ""}</div>
    ");


        var html = $@"
    <html>
    <head>
        <meta charset='utf-8'>
        <style>
            @font-face {{
                font-family: 'THSarabunNew';
                src: url('file:///{fontPath}') format('truetype');
                font-weight: normal;
                font-style: normal;
            }}
            body {{
                font-size: 22px;
                font-family: 'THSarabunNew', Arial, sans-serif;
            }}
            .t-18 {{ font-size: 1.5em; }}
            .t-18 {{ font-size: 1.7em; }}
            .t-22 {{ font-size: 1.9em; }}
                 .tab1 {{ text-indent: 48px;  word-break: break-all;  }}
        .tab2 {{ text-indent: 96px;  word-break: break-all; }}
        .tab3 {{ text-indent: 144px;  word-break: break-all; }}
        .tab4 {{ text-indent: 192px;  word-break: break-all;}}
            .d-flex {{ display: flex; }}
            .w-100 {{ width: 100%; }}
            .w-40 {{ width: 40%; }}
            .w-50 {{ width: 50%; }}
            .w-60 {{ width: 60%; }}
            .text-center {{ text-align: center; }}
            .sign-single-right {{
                display: flex;
                flex-direction: column;
                position: relative;
                left: 20%;
            }}
            .sign-double {{ display: flex; }}
            .text-center-right-brake {{
                margin-left: 50%;
                word-break: break-all;
            }}
            .text-right {{ text-align: right; }}
            .contract, .section {{
                margin: 12px 0;
                line-height: 1.7;
            }}
            .section {{
                font-weight: bold;
                font-size: 1.2em;
                text-align: left;
                margin-top: 24px;
            }}
            .signature-table {{
                width: 100%;
                margin-top: 32px;
                border-collapse: collapse;
            }}
            .signature-table td {{
                padding: 16px;
                text-align: center;
                vertical-align: top;
                font-size: 1.1em;
            }}
   .table {{ width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 28pt; }}
        .table th, .table td {{ border: 1px solid #000; padding: 8px; }}

        </style>
    </head>
    <body>
        {htmlBody}
    </body>
    </html>
    ";

        var doc = new HtmlToPdfDocument()
        {
            GlobalSettings = {
            PaperSize = PaperKind.A4,
            Orientation = Orientation.Portrait,
            Margins = new MarginSettings
            {
                Top = 20,
                Bottom = 20,
                Left = 20,
                Right = 20
            }
        },
            Objects = {
            new ObjectSettings() {
                HtmlContent = html,
                FooterSettings = new FooterSettings
                {
                    FontName = "THSarabunNew",
                    FontSize = 6,
                    Line = false,
                    Center = "[page] / [toPage]"
                }
            }
        }
        };

        var pdfBytes = _pdfConverter.Convert(doc);
        return pdfBytes;
    }
}
