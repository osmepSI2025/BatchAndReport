using BatchAndReport.DAO;
using BatchAndReport.Models;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using Microsoft.AspNetCore.Components.Web;
using PuppeteerSharp;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
public class WordSME_ReportService
{
    private readonly WordServiceSetting _w;
    private readonly Econtract_Report_GADAO _e;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    private readonly SmeDAO _smeDao;
    public WordSME_ReportService(WordServiceSetting ws, Econtract_Report_GADAO e
        , IConverter pdfConverter
        , SmeDAO smeDao
        )
    {
        _w = ws;
        _e = e;
        _pdfConverter = pdfConverter; // กำหนดค่า DI สำหรับ PDF Converter
        _smeDao = smeDao;
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
<div style='font-size:18pt;' ><b>ภาพรวมโครงการและงบประมาณเพื่อการส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (SME)</br>
ภายใต้แผนปฏิบัติการส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ประจำปี พ.ศ. {year}</b></div>
<hr>
<table style='width: 100%; border:1px solid #000; border-collapse:collapse; font-size:18pt;'>
    <tr style='font-weight:bold;background:#f0f0f0; font-size:18pt;'>
        <th style='border:1px solid #000; font-size:18pt;'>ประเด็นการส่งเสริม SME</th>
        <th style='border:1px solid #000; font-size:18pt;'>จำนวนโครงการ</th>
        <th style='border:1px solid #000; font-size:18pt;'>งบประมาณ (ล้านบาท)</th>
    </tr>
");

        int i = 1;
        foreach (var row in projects)
        {
            htmlBody.Append($@"
    <tr style='font-size:18pt;'>
        <td style='border:1px solid #000; font-size:18pt;' >ประเด็นที่ {row.IssueName ?? ""}</td>
        <td style='border:1px solid #000; font-size:18pt;'>{row.ProjectCount?.ToString("N0", culture) ?? "0"}</td>
        <td class='text-right' style='border:1px solid #000; font-size:18pt;'>{(row.Budget.GetValueOrDefault() / 1_000_000).ToString("N4", culture)}</td>
    </tr>
");
            i++;
        }

        htmlBody.Append($@"
    <tr style='font-weight:bold;background:#f0f0f0;' >
        <td style='border:1px solid #000; font-size:18pt;'>รวมทั้งหมด</td>
        <td style='border:1px solid #000;font-size:18pt;'>{totalProjects.ToString("N0", culture)}</td>
        <td class='text-right' style='border:1px solid #000;'>{(totalBudget / 1_000_000).ToString("N4", culture)}</td>
    </tr>
</table>
<br>
<div style='font-size:18pt;>โดยมีหน่วยงานทั้งหมด {strategyList.Where(p => !string.IsNullOrEmpty(p.Ministry_Id)).Select(p => p.Ministry_Id).Distinct().Count()} กระทรวง {strategyList.Where(p => !string.IsNullOrEmpty(p.DepartmentCode)).Select(p => p.DepartmentCode).Distinct().Count()} หน่วยงาน</div>
<br>
");
        // ------------------ Part 2: Strategy Detail ------------------
        var grouped = strategyList.GroupBy(p => p.Topic).Distinct().ToList();
        int topicIndex = 1;
        foreach (var topicGroup in grouped)
        {
            var TopictotalProject = topicGroup.Count();
            var TopicsumBudget = topicGroup.Sum(p => p.BudgetAmount);
            htmlBody.Append($@"<div style='font-size:18pt;'><b>ประเด็นการส่งเสริมที่  {topicGroup.Key}</b></div>");
            htmlBody.Append($@"<div style='font-size:18pt;'>จำนวน {TopictotalProject} โครงการ งบประมาณ {TopicsumBudget:N2} ล้านบาท</div>");
            htmlBody.Append($@"</br>");

            var strategyGrouped = topicGroup.GroupBy(p => p.StrategyDesc).ToList();
            int strategyIndex = 1;
            foreach (var strategyGroup in strategyGrouped)
            {
                var totalProject = strategyGroup.Count();
                var sumBudget = strategyGroup.Sum(p => p.BudgetAmount);
              
           

                htmlBody.Append($@"<div style='font-size:18pt;'><b>กลยุทธ์ที่  {strategyGroup.Key}</b></div>");        
                htmlBody.Append($@"<div style='font-size:18pt;'>จำนวน {totalProject} โครงการ งบประมาณ {sumBudget:N2} ล้านบาท</div>");

                // Table
                htmlBody.Append($@"
    <table style='width: 100%; border:1px solid #000; border-collapse:collapse;'>
        <tr>
            <th style='border:1px solid #000;font-size:18pt;' >หน่วยงาน/โครงการ</th>
            <th style='border:1px solid #000;font-size:18pt;' >งบประมาณ</th>
        </tr>
");

                var deptGrouped = strategyGroup.GroupBy(p => new { p.DepartmentCode, p.Department }).ToList();
                int projectIndex = 1;
                foreach (var deptGroup in deptGrouped)
                {
                    var deptTotal = deptGroup.Sum(p => p.BudgetAmount);
                    htmlBody.Append($@"
        <tr style='font-weight:bold;background:#e0e0e0;'>
            <td style='border:1px solid #000;font-size:18pt;' >{deptGroup.Key.Department}</td>
            <td style='border:1px solid #000;font-size:18pt; text-align:right;' >{deptTotal:N2}</td>
        </tr>
    ");
                    foreach (var proj in deptGroup)
                    {
                        htmlBody.Append($@"
            <tr>
                <td style='border:1px solid #000;font-size:18pt; padding-left:32px;' >{projectIndex}. {proj.ProjectName}</td>
                <td style='border:1px solid #000;font-size:18pt; text-align:right;' >{proj.BudgetAmount:N2}</td>
            </tr>
        ");
                        projectIndex++;
                    }
                }

                htmlBody.Append("</table><br>");
                strategyIndex++;
            }
            topicIndex++;
            //htmlBody.Append("<div style='page-break-after:always;'></div>");
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
        // Read CSS file content
        var cssPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "css", "contract.css");
        string contractCss = "";
        if (File.Exists(cssPath))
        {
            contractCss = File.ReadAllText(cssPath, Encoding.UTF8);
        }
        // Read logo and convert to Base64 (if needed in HTML)
        var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
        string logoBase64 = "";
        if (System.IO.File.Exists(logoPath))
        {
            var bytes = await System.IO.File.ReadAllBytesAsync(logoPath);
            logoBase64 = Convert.ToBase64String(bytes);
        }

        // Absolute font path for PDF rendering
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabun.ttf");
        string fontBase64 = "";
        if (File.Exists(fontPath))
        {
            var bytes = File.ReadAllBytes(fontPath);
            fontBase64 = Convert.ToBase64String(bytes);
        }
        // Build HTML content
        var htmlBody = new StringBuilder();



        // Heading
        htmlBody.Append($@"
        <div class='t-22 text-center'><b>แบบฟอร์มการจัดทำข้อเสนอ ประจำปีงบประมาณ {model.FiscalYearDesc}</b></div>
        <br>
        <div class='t-12'>กระทรวง : {model.MinistryName}</div>
        <div class='t-12'>หน่วยงาน : {model.DepartmentName}</div>
        <div class='t-12'>ชื่อกิจกรรม : {model.ProjectName}</div>
        <div class='t-12'>งบประมาณ : {model.BudgetAmount:N0}</div>
     

    
    ");

        if (model.IS_BUDGET_USED_FLAG == "Y")
        {
            htmlBody.Append($@"
       <div class='t-12'> ใช้งบประมาณ :  {model.BUDGET_SOURCE_NAME}</div>
     
    ");
        }
        else
        {
            htmlBody.Append($@"
  
        <div class='t-12'> ไม่ใช้งบประมาณ</div>
   
   
    ");
        }
        if (model.ProjectStatus == "01")
        {
            htmlBody.Append(@" 
        <div class='t-12'>สถานภาพโครงการ : โครงการใหม่ </div>
    ");
                
        }
        else if (model.ProjectStatus == "02")
        {
            htmlBody.Append(@" 
        <div class='t-12'>สถานภาพโครงการ :โครงการต่อเนื่อง </div>
    ");
        }
        else if (model.ProjectStatus == "03")
        {
            htmlBody.Append(@" 
        <div class='t-12'>สถานภาพโครงการ :โครงการเดิม </div>
    ");
        }
        else if (model.ProjectStatus == "04") 
        {

            htmlBody.Append(@" 
        <div class='t-12'>สถานภาพโครงการ : โครงการ Flagship</div>
    ");
        }


        var OwnerModel = await _smeDao.GetOwnerAndContactAsync(model.ProjectCode);
        // Responsible Table
        htmlBody.Append($@"
<table style='width:100%; border-collapse:collapse; border:1px solid #000;'>
    <tr style='background:#e0e0e0; font-weight:bold;'>
                <td style='border:1px solid #000;'></td>
                <td style='border:1px solid #000;text-align:center;'>ผู้รับผิดชอบโครงการ</td>
                <td style='border:1px solid #000;text-align:center;'>ผู้ประสานงาน</td>
            </tr>
            <tr>
                <td style='border:1px solid #000;text-align:center;'>ชื่อ-นามสกุล</td>
                <td style='border:1px solid #000;'>{OwnerModel.ToList()[0].OwnerName??""}</td>
                <td style='border:1px solid #000;'>{OwnerModel.ToList()[1].ContactName ?? ""}</td>
            </tr>
            <tr>
                <td style='border:1px solid #000;text-align:center;'>ตำแหน่ง</td>
                <td style='border:1px solid #000;'>{OwnerModel.ToList()[0].OwnerPosition}</td>
                <td style='border:1px solid #000;'>{OwnerModel.ToList()[1].ContactPosition ?? ""}</td>
            </tr>
            <tr>
                <td style='border:1px solid #000;text-align:center;'>โทรศัพท์</td>
                <td style='border:1px solid #000;'>{OwnerModel.ToList()[0].OwnerPhone ?? ""}</td>
                <td style='border:1px solid #000;'>{OwnerModel.ToList()[1].ContactPhone ?? ""}</td>
            </tr>
            <tr>
                <td style='border:1px solid #000;text-align:center;'>มือถือ</td>
                <td style='border:1px solid #000;'>{OwnerModel.ToList()[0].OwnerMobile ?? ""}</td>
                <td style='border:1px solid #000;'>{OwnerModel.ToList()[1].ContactMobile ?? ""}</td>
            </tr>
            <tr>
                <td style='border:1px solid #000;text-align:center;'>Email</td>
                <td style='border:1px solid #000;'>{OwnerModel.ToList()[0].OwnerEmail ?? ""}</td>
                <td style='border:1px solid #000;'>{OwnerModel.ToList()[1].ContactEmail ?? ""}</td>
            </tr>
            <tr>
                <td style='border:1px solid #000;text-align:center;'>Line ID</td>
                <td style='border:1px solid #000;'>{OwnerModel.ToList()[0].OwnerLineId ?? ""}</td>
                <td style='border:1px solid #000;'>{OwnerModel.ToList()[1].ContactLineId ?? ""}</td>
            </tr>
        </table>
        <br>
    ");

        if (model.SME_ISSUE_CODE == "SI01")
        {
            htmlBody.Append(@" 
           <div  class='tab1 t-12'><b>X Digital □ Environment/Green □ Social □ Governance □ Soft power</b></div>
    ");

        }
        else if (model.ProjectStatus == "SI02")
        {
            htmlBody.Append(@" 
          <div  class='tab1 t-12'><b>□ Digital X Environment/Green □ Social □ Governance □ Soft power</b></div>
    ");
        }
        else if (model.ProjectStatus == "SI03")
        {
            htmlBody.Append(@" 
         <div  class='tab1 t-12'><b>□ Digital □ Environment/Green X Social □ Governance □ Soft power</b></div>
    ");
        }
        else if (model.ProjectStatus == "SI04")
        {

            htmlBody.Append(@" 
          <div  class='tab1 t-12'><b>□ Digital □ Environment/Green □ Social X Governance □ Soft power</b></div>
    ");
        }
        else if (model.ProjectStatus == "SI054")
        {

            htmlBody.Append(@" 
          <div  class='tab1 t-12'><b>□ Digital □ Environment/Green □ Social □ Governance X Soft power</b></div>
    ");
        }

        htmlBody.Append($@"
        <div  class='tab1 t-12'><b>ประเด็นสำคัญในการส่งเสริม SME ปี พ.ศ.{model.FiscalYearDesc}</b></div>  ");
        if (model.SME_ISSUE_CODE == "SI01")
        {
            htmlBody.Append(@" 
           <div  class='tab1 t-12'><b> Digital </b></div>
    ");

        }
        else if (model.ProjectStatus == "SI02")
        {
            htmlBody.Append(@" 
          <div  class='tab1 t-12'><b> Environment/Green </b></div>
    ");
        }
        else if (model.ProjectStatus == "SI03")
        {
            htmlBody.Append(@" 
         <div  class='tab1 t-12'><b> Social</b></div>
    ");
        }
        else if (model.ProjectStatus == "SI04")
        {

            htmlBody.Append(@" 
          <div  class='tab1 t-12'><b>Governance </b></div>
    ");
        }
        else if (model.ProjectStatus == "SI054")
        {

            htmlBody.Append(@" 
          <div  class='tab1 t-12'><b> Soft power</b></div>
    ");
        }
        htmlBody.Append($@"<div  class='tab1 t-12'>ประเด็นสำคัญในการส่งเสริม SME ปี พ.ศ.{model.FiscalYearDesc} ประเด็นการส่งเสริม/กลยุทธ์ที่สอดคล้องกับแผนปฎิบัติการส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อมประจำปีงบประมาณ (เลือกเพียง 1 ประเด็นการส่งเสริม 1 กลยุทธ์ ต่อ 1 โครงการ)</div>
    ");

        // Strategy Table
        htmlBody.Append($@"
        <table style='width:100%; border-collapse:collapse;'>
            <tr class='t-12' style='background:#e0e0e0; font-weight:bold;'>
                <td style='border:1px solid #000;text-align:center;'>ประเด็นการส่งเสริม</td>
                <td style='border:1px solid #000;text-align:center;'>กลยุทธ์</td>
            </tr>
            <tr class='t-12' >
                <td style='border:1px solid #000;text-align:center;'>{model.Topic}</td>
                <td style='border:1px solid #000;text-align:center;'>{model.STRATEGY_DESC}</td>
            </tr>
    ");
        //if (model.Strategies != null)
        //{
        //    int index = 1;
        //    foreach (var strategy in model.Strategies)
        //    {
        //        htmlBody.Append($@"
        //        <tr>
        //            <td style='border:1px solid #000;'>□ {strategy.StrategyId}</td>
        //           <td style='border:1px solid #000;'>□ {index++} {strategy.StrategyDesc}</td>
        //        </tr>
        //    ");
        //    }
        //}
        htmlBody.Append("</table><br>");

        htmlBody.Append($@"
        <div  class='tab1 t-12'><b>ความสำคัญของโครงการ/หลักการและเหตุผล :</b></div>
        <div  class='tab1 t-12'>{model.ProjectRationale ?? ""}</div>
        <div  class='tab1 t-12'><b>วัตถุประสงค์ของโครงการ :</b></div>
        <div  class='tab1 t-12'>{model.ProjectObjective ?? ""}</div>
        <div  class='tab1 t-12'><b>กลุ่มเป้าหมาย</b></div>");

        if (model.TargetGroup != null)
        {

            //var docx = XDocument.Parse(model.TargetGroup);
            //// var docx = XDocument.Parse("G01");
            // var lookupCodes = docx.Descendants("field")
            //     .Where(f => (string)f.Attribute("name") == "LookupCode")
            //     .Select(f => (string)f.Element("value"))
            //     .ToList();
            var lookupCodes = model.TargetGroup?.Split(';', StringSplitOptions.RemoveEmptyEntries);
            foreach (var code in lookupCodes)
            {
                var lookup = await _smeDao.GetLookupAsync("TARGET_GROUP", code);
                if (lookup != null && !string.IsNullOrEmpty(lookup.LookupValue))
                {
                    htmlBody.Append($@"<div class='tab2'> {lookup.LookupValue}</div>");
                }
            }

        }


        htmlBody.Append($@"<div  class='tab1 t-12'><b>รายละเอียดแผนการดำเนินงาน/กิจกรรม...</b></div>
        <div  class='tab1 t-12'>{model.Activities ?? ""}</div>
        <div  class='tab1 t-12'><b>จุดเด่นของโครงการ :</b></div>
        <div  class='tab1 t-12'>{model.ProjectFocus ?? ""}</div>
        <div  class='tab1 t-12'><b>พื้นที่ดำเนินการ :</b></div>
        <div  class='tab1 t-12'>{string.Join(", ", model.OperationArea ?? new List<string>())}</div>
        <div  class='tab1 t-12'><b>สาขาเป้าหมาย :{model.TARGET_BRANCH_LIST}</b></div>
        <div  class='tab1 t-12'>{string.Join(", ", model.IndustrySector ?? new List<string>())}</div>
        <div  class='tab1 t-12'><b>การพัฒนา 11 อุตสาหกรรม Soft Power :</b></div>

        <div  class='tab1 t-12'><b>ระยะเวลาในการดำเนินโครงการ :{model.DaysDiff ?? ""} วัน</b></div>
 
        <div  class='tab1 t-12'><b>หน่วยงานที่ร่วมบูรณาการ...</b></div>
        <div  class='tab1 t-12'>{model.Partner_Name} ทำหน้าที่ {model.RoleDescription}</div>
        <div  class='tab1 t-12'><b>ตัวชี้วัดที่สำคัญ...</b></div>
    ");




        // Output Indicators Table
        htmlBody.Append($@"
   <table style='width:100%; border-collapse:collapse;'>
            <tr class='tab1 t-12' style='background:#e0e0e0; font-weight:bold;'>
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
                <tr class='tab1 t-12'>
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
                <tr >
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
        <div class='t-12'>ข้อมูลอื่นๆ เพิ่มเติม...</div>
        <div class='tab1 t-12'>{model.ProjectFocus ?? ""}</div>
    ");


        var html = $@"
    <html>
    <head>
        <meta charset='utf-8'>
        <style>
                 @font-face {{
            font-family: 'TH Sarabun PSK';
              src: url('data:font/truetype;charset=utf-8;base64,{fontBase64}') format('truetype');
            font-weight: normal;
            font-style: normal;
        }}
{contractCss}
        </style>
    </head>
    <body>
        {htmlBody}
    </body>
    </html>
    ";
        await new BrowserFetcher().DownloadAsync();
        await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
        await using var page = await browser.NewPageAsync();

        await page.SetContentAsync(html);

        var pdfOptions = new PdfOptions
        {
            Format = PuppeteerSharp.Media.PaperFormat.A4,

            Landscape = false,
            MarginOptions = new PuppeteerSharp.Media.MarginOptions
            {
                Top = "20mm",
                Bottom = "20mm",
                Left = "20mm",
                Right = "20mm"
            },
            PrintBackground = true

        };

        var pdfBytes = await page.PdfDataAsync(pdfOptions);  
      
        return pdfBytes;
    }

    public async Task<string> ExportSMEProjectDetail_HTML(
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
        <div class='t-12'>กระทรวง : {model.MinistryName}</div>
        <div class='t-12'>หน่วยงาน : {model.DepartmentName}</div>
        <div class='t-12'>ชื่อกิจกรรม : {model.ProjectName}</div>
        <div class='t-12'>งบประมาณ : {model.BudgetAmount:N0}</div>
     

    
    ");

        if (model.IS_BUDGET_USED_FLAG == "Y")
        {
            htmlBody.Append($@"
       <div class='t-12'> ใช้งบประมาณ :  {model.BUDGET_SOURCE_NAME}</div>
     
    ");
        }
        else
        {
            htmlBody.Append($@"
  
        <div class='t-12'> ไม่ใช้งบประมาณ</div>
   
   
    ");
        }
        if (model.ProjectStatus == "01")
        {
            htmlBody.Append(@" 
        <div class='t-12'>สถานภาพโครงการ : โครงการใหม่ </div>
    ");

        }
        else if (model.ProjectStatus == "02")
        {
            htmlBody.Append(@" 
        <div class='t-12'>สถานภาพโครงการ :โครงการต่อเนื่อง </div>
    ");
        }
        else if (model.ProjectStatus == "03")
        {
            htmlBody.Append(@" 
        <div class='t-12'>สถานภาพโครงการ :โครงการเดิม </div>
    ");
        }
        else if (model.ProjectStatus == "04")
        {

            htmlBody.Append(@" 
        <div class='t-12'>สถานภาพโครงการ : โครงการ Flagship</div>
    ");
        }


        var OwnerModel = await _smeDao.GetOwnerAndContactAsync(model.ProjectCode);
        // Responsible Table
        htmlBody.Append($@"
<table style='width:100%; border-collapse:collapse; border:1px solid #000;'>
    <tr class='t-12' style='background:#e0e0e0; font-weight:bold;'>
                <td class='t-12' style='border:1px solid #000;'></td>
                <td class='t-12' style='border:1px solid #000;text-align:center;'>ผู้รับผิดชอบโครงการ</td>
                <td class='t-12' style='border:1px solid #000;text-align:center;'>ผู้ประสานงาน</td>
            </tr>
            <tr>
                <td class='t-12' style='border:1px solid #000;text-align:center;'>ชื่อ-นามสกุล</td>
                <td class='t-12' style='border:1px solid #000;'>{OwnerModel.ToList()[0].OwnerName ?? ""}</td>
                <td class='t-12' style='border:1px solid #000;'>{OwnerModel.ToList()[1].ContactName ?? ""}</td>
            </tr>
            <tr>
                <td class='t-12' style='border:1px solid #000;text-align:center;'>ตำแหน่ง</td>
                <td class='t-12' style='border:1px solid #000;'>{OwnerModel.ToList()[0].OwnerPosition}</td>
                <td class='t-12' style='border:1px solid #000;'>{OwnerModel.ToList()[1].ContactPosition ?? ""}</td>
            </tr>
            <tr>
                <td class='t-12' style='border:1px solid #000;text-align:center;'>โทรศัพท์</td>
                <td class='t-12' style='border:1px solid #000;'>{OwnerModel.ToList()[0].OwnerPhone ?? ""}</td>
                <td class='t-12' style='border:1px solid #000;'>{OwnerModel.ToList()[1].ContactPhone ?? ""}</td>
            </tr>
            <tr>
                <td class='t-12' style='border:1px solid #000;text-align:center;'>มือถือ</td>
                <td class='t-12' style='border:1px solid #000;'>{OwnerModel.ToList()[0].OwnerMobile ?? ""}</td>
                <td class='t-12' style='border:1px solid #000;'>{OwnerModel.ToList()[1].ContactMobile ?? ""}</td>
            </tr>
            <tr>
                <td class='t-12' style='border:1px solid #000;text-align:center;'>Email</td>
                <td class='t-12' style='border:1px solid #000;'>{OwnerModel.ToList()[0].OwnerEmail ?? ""}</td>
                <td class='t-12' style='border:1px solid #000;'>{OwnerModel.ToList()[1].ContactEmail ?? ""}</td>
            </tr>
            <tr>
                <td class='t-12' style='border:1px solid #000;text-align:center;'>Line ID</td>
                <td class='t-12' style='border:1px solid #000;'>{OwnerModel.ToList()[0].OwnerLineId ?? ""}</td>
                <td class='t-12' style='border:1px solid #000;'>{OwnerModel.ToList()[1].ContactLineId ?? ""}</td>
            </tr>
        </table>
        <br>
    ");

        if (model.SME_ISSUE_CODE == "SI01")
        {
            htmlBody.Append(@" 
           <div  class='tab1 t-12'><b>X Digital □ Environment/Green □ Social □ Governance □ Soft power</b></div>
    ");

        }
        else if (model.ProjectStatus == "SI02")
        {
            htmlBody.Append(@" 
          <div  class='tab1 t-12'><b>□ Digital X Environment/Green □ Social □ Governance □ Soft power</b></div>
    ");
        }
        else if (model.ProjectStatus == "SI03")
        {
            htmlBody.Append(@" 
         <div  class='tab1 t-12'><b>□ Digital □ Environment/Green X Social □ Governance □ Soft power</b></div>
    ");
        }
        else if (model.ProjectStatus == "SI04")
        {

            htmlBody.Append(@" 
          <div  class='tab1 t-12'><b>□ Digital □ Environment/Green □ Social X Governance □ Soft power</b></div>
    ");
        }
        else if (model.ProjectStatus == "SI054")
        {

            htmlBody.Append(@" 
          <div  class='tab1 t-12'><b>□ Digital □ Environment/Green □ Social □ Governance X Soft power</b></div>
    ");
        }

        htmlBody.Append($@"
        <div  class='tab1 t-12'><b>ประเด็นสำคัญในการส่งเสริม SME ปี พ.ศ.{model.FiscalYearDesc}</b></div>  ");
        if (model.SME_ISSUE_CODE == "SI01")
        {
            htmlBody.Append(@" 
           <div  class='tab1 t-12'><b> Digital </b></div>
    ");

        }
        else if (model.ProjectStatus == "SI02")
        {
            htmlBody.Append(@" 
          <div  class='tab1 t-12'><b> Environment/Green </b></div>
    ");
        }
        else if (model.ProjectStatus == "SI03")
        {
            htmlBody.Append(@" 
         <div  class='tab1 t-12'><b> Social</b></div>
    ");
        }
        else if (model.ProjectStatus == "SI04")
        {

            htmlBody.Append(@" 
          <div  class='tab1 t-12'><b>Governance </b></div>
    ");
        }
        else if (model.ProjectStatus == "SI054")
        {

            htmlBody.Append(@" 
          <div  class='tab1 t-12'><b> Soft power</b></div>
    ");
        }
        htmlBody.Append($@"<div  class='tab1 t-12'>ประเด็นสำคัญในการส่งเสริม SME ปี พ.ศ.{model.FiscalYearDesc} ประเด็นการส่งเสริม/กลยุทธ์ที่สอดคล้องกับแผนปฎิบัติการส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อมประจำปีงบประมาณ (เลือกเพียง 1 ประเด็นการส่งเสริม 1 กลยุทธ์ ต่อ 1 โครงการ)</div>
    ");

        // Strategy Table
        htmlBody.Append($@"
        <table  style='width:100%; border-collapse:collapse;'>
            <tr style='background:#e0e0e0; font-weight:bold;'>
                <td style='border:1px solid #000;text-align:center;'>ประเด็นการส่งเสริม</td>
                <td style='border:1px solid #000;text-align:center;'>กลยุทธ์</td>
            </tr>
            <tr >
                <td style='border:1px solid #000;text-align:center;'>{model.Topic}</td>
                <td style='border:1px solid #000;text-align:center;'>{model.STRATEGY_DESC}</td>
            </tr>
    ");

        htmlBody.Append("</table><br>");

        htmlBody.Append($@"
        <div  class='tab1'><b>ความสำคัญของโครงการ/หลักการและเหตุผล :</b></div>
        <div  class='tab1'>{model.ProjectRationale ?? ""}</div>
        <div  class='tab1'><b>วัตถุประสงค์ของโครงการ :</b></div>
        <div  class='tab1'>{model.ProjectObjective ?? ""}</div>
        <div  class='tab1'><b>กลุ่มเป้าหมาย</b></div>");

        if (model.TargetGroup != null)
        {

            //var docx = XDocument.Parse(model.TargetGroup);
            //// var docx = XDocument.Parse("G01");
            // var lookupCodes = docx.Descendants("field")
            //     .Where(f => (string)f.Attribute("name") == "LookupCode")
            //     .Select(f => (string)f.Element("value"))
            //     .ToList();
            var lookupCodes = model.TargetGroup?.Split(';', StringSplitOptions.RemoveEmptyEntries);
            foreach (var code in lookupCodes)
            {
                var lookup = await _smeDao.GetLookupAsync("TARGET_GROUP", code);
                if (lookup != null && !string.IsNullOrEmpty(lookup.LookupValue))
                {
                    htmlBody.Append($@"<div class='tab2'> {lookup.LookupValue}</div>");
                }
            }

        }


        htmlBody.Append($@"<div  class='tab1'><b>รายละเอียดแผนการดำเนินงาน/กิจกรรม...</b></div>
        <div  class='tab1'>{model.Activities ?? ""}</div>
        <div  class='tab1'><b>จุดเด่นของโครงการ :</b></div>
        <div  class='tab1'>{model.ProjectFocus ?? ""}</div>
        <div  class='tab1'><b>พื้นที่ดำเนินการ :</b></div>
        <div  class='tab1'>{string.Join(", ", model.OperationArea ?? new List<string>())}</div>
        <div  class='tab1'><b>สาขาเป้าหมาย :{model.TARGET_BRANCH_LIST}</b></div>
        <div  class='tab1'>{string.Join(", ", model.IndustrySector ?? new List<string>())}</div>
        <div  class='tab1'><b>การพัฒนา 11 อุตสาหกรรม Soft Power :</b></div>

        <div  class='tab1'><b>ระยะเวลาในการดำเนินโครงการ :{model.DaysDiff ?? ""} วัน</b></div>
 
        <div  class='tab1'><b>หน่วยงานที่ร่วมบูรณาการ...</b></div>
        <div  class='tab1'>{model.Partner_Name} ทำหน้าที่ {model.RoleDescription}</div>
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
        <div class='t-12'>ข้อมูลอื่นๆ เพิ่มเติม...</div>
        <div class='t-12'>{model.AdditionalNotes ?? ""}</div>
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

        return html;
    }
    public byte[] ConvertHtmlToWord(string html)
    {
        using (var mem = new MemoryStream())
        {
            using (var wordDoc = WordprocessingDocument.Create(mem, WordprocessingDocumentType.Document, true))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // Use HtmlToOpenXml to convert HTML to OpenXML
                var converter = new HtmlToOpenXml.HtmlConverter(mainPart);
                var paragraphs = converter.Parse(html);
                mainPart.Document.Body.Append(paragraphs);

                // wordDoc.Close(); // Remove this line, not needed
            }
            return mem.ToArray();
        }
    }

    
}
