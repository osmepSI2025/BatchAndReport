using BatchAndReport.DAO;
using BatchAndReport.Models;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using System.Text;
public class WordWorkFlow_annualProcessReviewService
{
    private readonly WordServiceSetting _w;
    private readonly Econtract_Report_GADAO _e;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    public WordWorkFlow_annualProcessReviewService(WordServiceSetting ws, Econtract_Report_GADAO e
        , IConverter pdfConverter
        )
    {
        _w = ws;
        _e = e;
        _pdfConverter = pdfConverter; // กำหนดค่า DI สำหรับ PDF Converter
    }


    public async Task<byte[]> GenAnnualWorkProcesses_HtmlToPDF(WFProcessDetailModels detail)
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
        <div class='text-center t-18'>
            <b>การทบทวนกระบวนการของ {detail.BusinessUnitOwner} ประจำปี {detail.FiscalYear}</b>
        </div>
    ");

        // Numbered review details
        if (detail.ReviewDetails?.Length > 0)
        {
            htmlBody.Append("</br>");
            htmlBody.Append("<div  class=' t-18'>ความเป็นมา</div>");
            if (!string.IsNullOrEmpty(detail.PROCESS_BACKGROUND))
            {
                var lines = detail.PROCESS_BACKGROUND
                    .Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)
                    .Select(line => $"<div class='tab1 t-18'>{System.Net.WebUtility.HtmlEncode(line)}</div>");
                htmlBody.Append(string.Join("", lines));
            }

            htmlBody.Append("</br>");
            htmlBody.Append("<div  class='t-18' >รายละเอียดประเด็นการทบทวน</div>");
            htmlBody.Append("<ol   style='margin-left:32px;'>");
            foreach (var item in detail.ReviewDetails)
                htmlBody.Append($"<li  class='t-18'>{System.Net.WebUtility.HtmlEncode(item)}</li>");
            htmlBody.Append("</ol>");
        }

        // Section heading
        htmlBody.Append($@"
        <div  class='t-18' >
            การทบทวนกระบวนการของ {detail.BusinessUnitOwner} ประจำปี {detail.FiscalYear} ดังนี้
        </div>
    ");

        // Three-column table
        htmlBody.Append("<table class='w-100 t-18' border='1' cellpadding='6' style='border-collapse:collapse;margin-bottom:12px;'>");
        htmlBody.Append($@"
        <tr style='background:#DDEBF7;font-weight:bold; vertical-align: top; '>
            <td >กระบวนการ ปี {detail.FiscalYearPrevious} (เดิม)</td>
            <td >กระบวนการ ปี {detail.FiscalYear} (ทบทวน)</td>
            <td  >กระบวนการที่กำหนดกิจกรรมควบคุม (Control Activity) ส่งกรมบัญชีกลาง</td>
        </tr>
    ");
        int rowCount = Math.Max(
            Math.Max(detail.PrevProcesses?.Count ?? 0, detail.CurrentProcesses?.Count ?? 0),
            detail.ControlActivities?.Count ?? 0
        );
        for (int i = 0; i < rowCount; i++)
        {
            htmlBody.Append("<tr>");
            htmlBody.Append($"<td >{System.Net.WebUtility.HtmlEncode(detail.PrevProcesses?.ElementAtOrDefault(i) ?? "")}</td>");
            htmlBody.Append($"<td>{System.Net.WebUtility.HtmlEncode(detail.CurrentProcesses?.ElementAtOrDefault(i) ?? "")}</td>");
            htmlBody.Append($"<td >{System.Net.WebUtility.HtmlEncode(detail.ControlActivities?.ElementAtOrDefault(i) ?? "")}</td>");
            htmlBody.Append("</tr>");
        }
        htmlBody.Append("</table>");

        // Note
        htmlBody.Append("<div class='t-18' style='font-style:italic;margin-bottom:12px;'>หมายเหตุ: *ทบทวนตาม JD/ **ทบทวนตาม วค.2/***ทบทวนตามภารกิจงานปัจจุบัน</div>");

        // Workflow processes
        if (detail.WorkflowProcesses?.Count > 0)
        {
            htmlBody.Append("<div>กระบวนการที่จัดทำ Workflow เพิ่มเติม ได้แก่</div>");
            foreach (var wf in detail.WorkflowProcesses)
                htmlBody.Append($"<div style='margin-left:32px;'>• {System.Net.WebUtility.HtmlEncode(wf)}</div>");
        }

        // Comments
        htmlBody.Append("<div class='t-18' >ความคิดเห็น</div>");
        htmlBody.Append("<div class='t-18' style='margin-left:32px;'>☐ เห็นชอบการปรับปรุง</div>");
        htmlBody.Append("<div class='t-18' style='margin-left:32px;'>☐ มีความเห็นเพิ่มเติม</div>");

        // Approve remarks
        if (detail.ApproveRemarks?.Length > 0)
        {
            foreach (var r in detail.ApproveRemarks)
                htmlBody.Append($"<div class='t-18' style='margin-left:32px;'>{System.Net.WebUtility.HtmlEncode(r)}</div>");
        }

        // Signature section
        htmlBody.Append(@"
        <table class='signature-table t-18'>
            <tr>
                <td>
                    <div>ลงชื่อ....................................................</div>
                    <div>(" + (detail.Approver1Name ?? "(ชื่อผู้ลงนาม 1)") + @")</div>
                    <div>" + (detail.Approver1Position ?? "ตำแหน่ง") + @"</div>
                    <div>วันที่ " + (detail.Approve1Date ?? "ไม่พบข้อมูล") + @"</div>
                </td>
                <td>
                    <div>ลงชื่อ....................................................</div>
                    <div>(" + (detail.Approver2Name ?? "(ชื่อผู้ลงนาม 2)") + @")</div>
                    <div>" + (detail.Approver2Position ?? "ตำแหน่ง") + @"</div>
                    <div>วันที่ " + (detail.Approve2Date ?? "ไม่พบข้อมูล") + @"</div>
                </td>
            </tr>
        </table>
    ");

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
    public async Task<string> GenAnnualWorkProcesses_Html(WFProcessDetailModels detail)
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
        <div class='text-center t-18'>
            <b>การทบทวนกระบวนการของ {detail.BusinessUnitOwner} ประจำปี {detail.FiscalYear}</b>
        </div>
    ");

        // Numbered review details
        if (detail.ReviewDetails?.Length > 0)
        {
            htmlBody.Append("</br>");
            htmlBody.Append("<div  class=' t-18'>ความเป็นมา</div>");
            if (!string.IsNullOrEmpty(detail.PROCESS_BACKGROUND))
            {
                var lines = detail.PROCESS_BACKGROUND
                    .Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)
                    .Select(line => $"<div class='tab1 t-18'>{System.Net.WebUtility.HtmlEncode(line)}</div>");
                htmlBody.Append(string.Join("", lines));
            }

            htmlBody.Append("</br>");
            htmlBody.Append("<div  class='t-18' >รายละเอียดประเด็นการทบทวน</div>");
            htmlBody.Append("<ol   style='margin-left:32px;'>");
            foreach (var item in detail.ReviewDetails)
                htmlBody.Append($"<li  class='t-18'>{System.Net.WebUtility.HtmlEncode(item)}</li>");
            htmlBody.Append("</ol>");
        }

        // Section heading
        htmlBody.Append($@"
        <div  class='t-18' >
            การทบทวนกระบวนการของ {detail.BusinessUnitOwner} ประจำปี {detail.FiscalYear} ดังนี้
        </div>
    ");

        // Three-column table
        htmlBody.Append("<table class='w-100 t-18' border='1' cellpadding='6' style='border-collapse:collapse;margin-bottom:12px;'>");
        htmlBody.Append($@"
        <tr style='background:#DDEBF7;font-weight:bold; vertical-align: top; '>
            <td >กระบวนการ ปี {detail.FiscalYearPrevious} (เดิม)</td>
            <td >กระบวนการ ปี {detail.FiscalYear} (ทบทวน)</td>
            <td  >กระบวนการที่กำหนดกิจกรรมควบคุม (Control Activity) ส่งกรมบัญชีกลาง</td>
        </tr>
    ");
        int rowCount = Math.Max(
            Math.Max(detail.PrevProcesses?.Count ?? 0, detail.CurrentProcesses?.Count ?? 0),
            detail.ControlActivities?.Count ?? 0
        );
        for (int i = 0; i < rowCount; i++)
        {
            htmlBody.Append("<tr>");
            htmlBody.Append($"<td >{System.Net.WebUtility.HtmlEncode(detail.PrevProcesses?.ElementAtOrDefault(i) ?? "")}</td>");
            htmlBody.Append($"<td>{System.Net.WebUtility.HtmlEncode(detail.CurrentProcesses?.ElementAtOrDefault(i) ?? "")}</td>");
            htmlBody.Append($"<td >{System.Net.WebUtility.HtmlEncode(detail.ControlActivities?.ElementAtOrDefault(i) ?? "")}</td>");
            htmlBody.Append("</tr>");
        }
        htmlBody.Append("</table>");

        // Note
        htmlBody.Append("<div class='t-18' style='font-style:italic;margin-bottom:12px;'>หมายเหตุ: *ทบทวนตาม JD/ **ทบทวนตาม วค.2/***ทบทวนตามภารกิจงานปัจจุบัน</div>");

        // Workflow processes
        if (detail.WorkflowProcesses?.Count > 0)
        {
            htmlBody.Append("<div>กระบวนการที่จัดทำ Workflow เพิ่มเติม ได้แก่</div>");
            foreach (var wf in detail.WorkflowProcesses)
                htmlBody.Append($"<div style='margin-left:32px;'>• {System.Net.WebUtility.HtmlEncode(wf)}</div>");
        }

        // Comments
        htmlBody.Append("<div class='t-18' >ความคิดเห็น</div>");
        htmlBody.Append("<div class='t-18' style='margin-left:32px;'>☐ เห็นชอบการปรับปรุง</div>");
        htmlBody.Append("<div class='t-18' style='margin-left:32px;'>☐ มีความเห็นเพิ่มเติม</div>");

        // Approve remarks
        if (detail.ApproveRemarks?.Length > 0)
        {
            foreach (var r in detail.ApproveRemarks)
                htmlBody.Append($"<div class='t-18' style='margin-left:32px;'>{System.Net.WebUtility.HtmlEncode(r)}</div>");
        }

        // Signature section
        htmlBody.Append(@"
        <table class='signature-table t-18'>
            <tr>
                <td>
                    <div>ลงชื่อ....................................................</div>
                    <div>(" + (detail.Approver1Name ?? "(ชื่อผู้ลงนาม 1)") + @")</div>
                    <div>" + (detail.Approver1Position ?? "ตำแหน่ง") + @"</div>
                    <div>วันที่ " + (detail.Approve1Date ?? "ไม่พบข้อมูล") + @"</div>
                </td>
                <td>
                    <div>ลงชื่อ....................................................</div>
                    <div>(" + (detail.Approver2Name ?? "(ชื่อผู้ลงนาม 2)") + @")</div>
                    <div>" + (detail.Approver2Position ?? "ตำแหน่ง") + @"</div>
                    <div>วันที่ " + (detail.Approve2Date ?? "ไม่พบข้อมูล") + @"</div>
                </td>
            </tr>
        </table>
    ");

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
        </style>
    </head>
    <body>
        {htmlBody}
    </body>
    </html>
    ";

        return html;
    }
    public async Task<byte[]> GenExportWorkProcesses_HtmlToPDF(WFProcessDetailModels detail)
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

        var htmlBody = new StringBuilder();

        // Header
        htmlBody.Append($@"
        <div class='text-center'>
            <b>การทบทวนกระบวนการของ {detail.BusinessUnitOwner} ประจำปี {detail.FiscalYear}</b>
        </div>
    ");

        // Core Process Table
        if (detail.CoreProcesses != null && detail.CoreProcesses.Count > 0)
        {
            htmlBody.Append("<table class='w-100' border='1' cellpadding='6' style='border-collapse:collapse;margin-bottom:12px; table-layout:fixed;'>");
            htmlBody.Append("<colgroup>");
            htmlBody.Append("<col style='width:20%;'/>"); // First column
            int coreCount = detail.CoreProcesses?.Count ?? 0;
            for (int i = 0; i < coreCount; i++)
                htmlBody.Append($"<col style='width:{80.0 / coreCount}%;'/>"); // Distribute remaining width equally
            htmlBody.Append("</colgroup>");
            // Row 1: กลุ่มกระบวนการหลัก + รหัส
            htmlBody.Append("<tr>");
            htmlBody.Append("<td rowspan='2' style='width:20%;font-weight:bold;background:#fff;'>กลุ่มกระบวนการหลัก<br/>(Core Process)</td>");
            foreach (var core in detail.CoreProcesses)
                htmlBody.Append($"<td class='t-18' style='background:#00C896;text-align:center;vertical-align:middle;'>{System.Net.WebUtility.HtmlEncode(core.ProcessGroupCode)}</td>");
            htmlBody.Append("</tr>");
            // Row 2: ชื่อกระบวนการ
            htmlBody.Append("<tr>");
            foreach (var core in detail.CoreProcesses)
                htmlBody.Append($"<td class='t-18' style='background:#00C896;text-align:center;vertical-align:middle;white-space:normal;word-break:break-word;'>{System.Net.WebUtility.HtmlEncode(core.ProcessGroupName)}</td>");
            htmlBody.Append("</tr>");
            htmlBody.Append("</table>");
        }

        // Supporting Process Table
        if (detail.SupportProcesses != null && detail.SupportProcesses.Count > 0)
        {
            htmlBody.Append("<table class='w-100 t-18' border='1' cellpadding='6' style='border-collapse:collapse;margin-bottom:12px; table-layout:fixed;'>");
            htmlBody.Append("<colgroup>");
            htmlBody.Append("<col style='width:20%;'/>"); // First column
            htmlBody.Append("<col style='width:10%;'/>"); // Code column
            htmlBody.Append("<col style='width:70%;'/>"); // Name column
            htmlBody.Append("</colgroup>");
            for (int i = 0; i < detail.SupportProcesses.Count; i++)
            {
                var support = detail.SupportProcesses[i];
                htmlBody.Append("<tr>");
                if (i == 0)
                {
                    htmlBody.Append($"<td rowspan='{detail.SupportProcesses.Count}' style='width:20%;font-weight:bold;'>กลุ่มกระบวนการสนับสนุน<br/>(Supporting Process)</td>");
                }
                htmlBody.Append($"<td class='t-18' style='background:#4CB1F0;text-align:center;width:10%;'>{System.Net.WebUtility.HtmlEncode(support.ProcessGroupCode)}</td>");
                htmlBody.Append($"<td class='t-18' style='background:#4CB1F0;text-align:left;width:70%;'>{System.Net.WebUtility.HtmlEncode(support.ProcessGroupName)}</td>");
                htmlBody.Append("</tr>");
            }
            htmlBody.Append("</table>");
        }

        // Section heading
    //    htmlBody.Append($@"
    //    <div>
    //        การทบทวนกระบวนการของ {detail.BusinessUnitOwner} ประจำปี {detail.FiscalYear} ดังนี้
    //    </div>
    //");

    //    // Three-column table
    //    int rowCount = Math.Max(
    //        Math.Max(detail.PrevProcesses?.Count ?? 0, detail.CurrentProcesses?.Count ?? 0),
    //        detail.ControlActivities?.Count ?? 0
    //    );
    //    htmlBody.Append("<table class='w-100' border='1' cellpadding='6' style='border-collapse:collapse;margin-bottom:12px;'>");
    //    htmlBody.Append($@"
    //    <tr style='background:#DDEBF7;font-weight:bold;text-align:center;'>
    //        <td>กระบวนการ ปี {detail.FiscalYearPrevious} (เดิม)</td>
    //        <td>กระบวนการ ปี {detail.FiscalYear} (ทบทวน)</td>
    //        <td>กระบวนการที่กำหนดกิจกรรมควบคุม (Control Activity) ส่งกรมบัญชีกลาง</td>
    //    </tr>
    //");
    //    for (int i = 0; i < rowCount; i++)
    //    {
    //        htmlBody.Append("<tr>");
    //        htmlBody.Append($"<td>{System.Net.WebUtility.HtmlEncode(detail.PrevProcesses?.ElementAtOrDefault(i) ?? "")}</td>");
    //        htmlBody.Append($"<td>{System.Net.WebUtility.HtmlEncode(detail.CurrentProcesses?.ElementAtOrDefault(i) ?? "")}</td>");
    //        htmlBody.Append($"<td>{System.Net.WebUtility.HtmlEncode(detail.ControlActivities?.ElementAtOrDefault(i) ?? "")}</td>");
    //        htmlBody.Append("</tr>");
    //    }
    //    htmlBody.Append("</table>");

        //// Note
        //htmlBody.Append("<div style='font-style:italic;margin-bottom:12px;'>หมายเหตุ: *ทบทวนตาม JD/ **ทบทวนตาม วค.2/***ทบทวนตามภารกิจงานปัจจุบัน</div>");

        //// Workflow processes
        //if (detail.WorkflowProcesses?.Count > 0)
        //{
        //    htmlBody.Append("<div>กระบวนการที่จัดทำ Workflow เพิ่มเติม ได้แก่</div>");
        //    foreach (var wf in detail.WorkflowProcesses)
        //        htmlBody.Append($"<div style='margin-left:32px;'>• {System.Net.WebUtility.HtmlEncode(wf)}</div>");
        //}

        //// Comments
        //htmlBody.Append("<div>ความคิดเห็น</div>");
        //htmlBody.Append("<div style='margin-left:32px;'>☐ เห็นชอบการปรับปรุง</div>");
        //htmlBody.Append("<div style='margin-left:32px;'>☐ มีความเห็นเพิ่มเติม</div>");

        //// Approve remarks
        //if (detail.ApproveRemarks?.Length > 0)
        //{
        //    foreach (var r in detail.ApproveRemarks)
        //        htmlBody.Append($"<div style='margin-left:32px;'>{System.Net.WebUtility.HtmlEncode(r)}</div>");
        //}

        // Signature section
    //    htmlBody.Append(@"
    //    <table class='signature-table'>
    //        <tr>
    //            <td>
    //                <div>ลงชื่อ....................................................</div>
    //                <div>(" + (detail.Approver1Name ?? "(ชื่อผู้ลงนาม 1)") + @")</div>
    //                <div>" + (detail.Approver1Position ?? "ตำแหน่ง") + @"</div>
    //                <div>วันที่ " + (detail.Approve1Date ?? "ไม่พบข้อมูล") + @"</div>
    //            </td>
    //            <td>
    //                <div>ลงชื่อ....................................................</div>
    //                <div>(" + (detail.Approver2Name ?? "(ชื่อผู้ลงนาม 2)") + @")</div>
    //                <div>" + (detail.Approver2Position ?? "ตำแหน่ง") + @"</div>
    //                <div>วันที่ " + (detail.Approve2Date ?? "ไม่พบข้อมูล") + @"</div>
    //            </td>
    //        </tr>
    //    </table>
    //");

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
    public async Task<string> GenExportWorkProcesses_Html(WFProcessDetailModels detail)
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

        var htmlBody = new StringBuilder();

        // Header
        htmlBody.Append($@"
        <div class='text-center'>
            <b>การทบทวนกระบวนการของ {detail.BusinessUnitOwner} ประจำปี {detail.FiscalYear}</b>
        </div>
    ");

        // Core Process Table
        if (detail.CoreProcesses != null && detail.CoreProcesses.Count > 0)
        {
            htmlBody.Append("<table class='w-100' border='1' cellpadding='6' style='border-collapse:collapse;margin-bottom:12px; table-layout:fixed;'>");
            htmlBody.Append("<colgroup>");
            htmlBody.Append("<col style='width:20%;'/>"); // First column
            int coreCount = detail.CoreProcesses?.Count ?? 0;
            for (int i = 0; i < coreCount; i++)
                htmlBody.Append($"<col style='width:{80.0 / coreCount}%;'/>"); // Distribute remaining width equally
            htmlBody.Append("</colgroup>");
            // Row 1: กลุ่มกระบวนการหลัก + รหัส
            htmlBody.Append("<tr>");
            htmlBody.Append("<td rowspan='2' style='width:20%;font-weight:bold;background:#fff;'>กลุ่มกระบวนการหลัก<br/>(Core Process)</td>");
            foreach (var core in detail.CoreProcesses)
                htmlBody.Append($"<td class='t-18' style='background:#00C896;text-align:center;vertical-align:middle;'>{System.Net.WebUtility.HtmlEncode(core.ProcessGroupCode)}</td>");
            htmlBody.Append("</tr>");
            // Row 2: ชื่อกระบวนการ
            htmlBody.Append("<tr>");
            foreach (var core in detail.CoreProcesses)
                htmlBody.Append($"<td class='t-18' style='background:#00C896;text-align:center;vertical-align:middle;white-space:normal;word-break:break-word;'>{System.Net.WebUtility.HtmlEncode(core.ProcessGroupName)}</td>");
            htmlBody.Append("</tr>");
            htmlBody.Append("</table>");
        }

        // Supporting Process Table
        if (detail.SupportProcesses != null && detail.SupportProcesses.Count > 0)
        {
            htmlBody.Append("<table class='w-100 t-18' border='1' cellpadding='6' style='border-collapse:collapse;margin-bottom:12px; table-layout:fixed;'>");
            htmlBody.Append("<colgroup>");
            htmlBody.Append("<col style='width:20%;'/>"); // First column
            htmlBody.Append("<col style='width:10%;'/>"); // Code column
            htmlBody.Append("<col style='width:70%;'/>"); // Name column
            htmlBody.Append("</colgroup>");
            for (int i = 0; i < detail.SupportProcesses.Count; i++)
            {
                var support = detail.SupportProcesses[i];
                htmlBody.Append("<tr>");
                if (i == 0)
                {
                    htmlBody.Append($"<td rowspan='{detail.SupportProcesses.Count}' style='width:20%;font-weight:bold;'>กลุ่มกระบวนการสนับสนุน<br/>(Supporting Process)</td>");
                }
                htmlBody.Append($"<td class='t-18' style='background:#4CB1F0;text-align:center;width:10%;'>{System.Net.WebUtility.HtmlEncode(support.ProcessGroupCode)}</td>");
                htmlBody.Append($"<td class='t-18' style='background:#4CB1F0;text-align:left;width:70%;'>{System.Net.WebUtility.HtmlEncode(support.ProcessGroupName)}</td>");
                htmlBody.Append("</tr>");
            }
            htmlBody.Append("</table>");
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

      
        return html;
    }
}
