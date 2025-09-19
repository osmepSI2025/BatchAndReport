using BatchAndReport.DAO;
using BatchAndReport.Models;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Signatures;
using System.Globalization;
using System.Text;
public class WordWorkFlow_annualProcessReviewService
{
    private readonly WordServiceSetting _w;
    private readonly Econtract_Report_GADAO _e;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    private readonly ILogger _logger;
    public WordWorkFlow_annualProcessReviewService(WordServiceSetting ws, Econtract_Report_GADAO e
        , IConverter pdfConverter
        , ILogger<WordWorkFlow_annualProcessReviewService> logger
        )
    {
        _w = ws;
        _e = e;
        _pdfConverter = pdfConverter; // กำหนดค่า DI สำหรับ PDF Converter
        _logger = logger;
    }


    public async Task<string> GenAnnualWorkProcesses_Html(WFProcessDetailModels detail,string flagSign )
    {
        var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
        string logoBase64 = "";
        if (System.IO.File.Exists(logoPath))
        {
            var bytes = await System.IO.File.ReadAllBytesAsync(logoPath);
            logoBase64 = Convert.ToBase64String(bytes);
        }

       // var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf");
        string fontBase64 = "";
        if (File.Exists(fontPath))
        {
            var bytes = File.ReadAllBytes(fontPath);
            fontBase64 = Convert.ToBase64String(bytes);
        }
        var htmlBody = new StringBuilder();
        var htmlTable = new StringBuilder();
        var htmlComment = new StringBuilder();
        var htmlDescript = new StringBuilder();
        var htmlSign = new StringBuilder();

        htmlTable.Append(@"
        <table border='1' cellpadding='6' style='border-collapse:collapse;width:100%;table-layout:fixed;'>
            <colgroup>
                <col style='width:33.33%;'/>
                <col style='width:33.33%;'/>
                <col style='width:33.33%;'/>
            </colgroup>
            <tr>
                <td class='t-12' style='font-weight:bold;background-color:#DDEBF7;text-align:center;'>กระบวนการ ปี " + detail.FiscalYearPrevious + @" (เดิม)</td>
                <td class='t-12' style='font-weight:bold;background-color:#DDEBF7;text-align:center;'>กระบวนการ ปี " + detail.FiscalYear + @" (ทบทวน)</td>
                <td class='t-12' style='font-weight:bold;background-color:#DDEBF7;text-align:center;'>กระบวนการที่กำหนด กิจกรรมควบคุม<br/>(Control Activity)<br/>ส่งกรมบัญชีกลาง</td>
            </tr>");
        int rowCount = Math.Max(
            Math.Max(detail.PrevProcesses?.Count ?? 0, detail.CurrentProcesses?.Count ?? 0),
            detail.ControlActivities?.Count ?? 0
        );
        for (int i = 0; i < rowCount; i++)
        {
            htmlTable.Append("<tr>");
            htmlTable.Append("<td class='t-12'>" + System.Net.WebUtility.HtmlEncode(detail.PrevProcesses?.ElementAtOrDefault(i) ?? "") + "</td>");
            htmlTable.Append("<td class='t-12'>" + System.Net.WebUtility.HtmlEncode(detail.CurrentProcesses?.ElementAtOrDefault(i) ?? "") + "</td>");
            htmlTable.Append("<td class='t-12'>" + System.Net.WebUtility.HtmlEncode(detail.ControlActivities?.ElementAtOrDefault(i) ?? "") + "</td>");
            htmlTable.Append("</tr>");
        }
        htmlTable.Append("</table>");





        #region ความคิดเห็น
        if (detail.wf_tasklist != null)
        {
            bool isApproveChecked = false;
            bool isCommentChecked = false;

            if (detail.wf_tasklist.STATUS == "ST0204")
            {
                if (!string.IsNullOrWhiteSpace(detail.commentDetial))
                {
                    isCommentChecked = true;
                }
                else
                {
                    isApproveChecked = true;
                }
            }

            htmlComment.Append(@"
<div class='comment-section'>
    <div class='t-12'>
        <input type='checkbox' style='transform:scale(1.3);margin-right:8px;' " + (isApproveChecked ? "checked" : "") + @" /> เห็นชอบการปรับปรุง
    </div>
    <div class='t-12'>
        <input type='checkbox' style='transform:scale(1.3);margin-right:8px;' " + (isCommentChecked ? "checked" : "") + @" /> มีความเห็นเพิ่มเติม
    </div>"
                + (!string.IsNullOrWhiteSpace(detail.commentDetial)
                    ? "<div class='tab2 t-12'>" + System.Net.WebUtility.HtmlEncode(detail.commentDetial) + "</div>"
                    : "")
                + @"
</div>
");
        }

        #endregion

        #region
        htmlDescript.Append("</br>");
        htmlDescript.Append("<div  class='t-12' >รายละเอียดประเด็นการทบทวน</div>");
        htmlDescript.Append("<ol   style='margin-left:32px;'>");
        foreach (var item in detail.ReviewDetails)
            htmlDescript.Append($"<li  class='t-12'>{System.Net.WebUtility.HtmlEncode(item)}</li>");
        htmlDescript.Append("</ol>");
        #endregion
        #region Signature Table


        if (detail.approvelist != null && detail.approvelist.Count > 0)
        {
            string noSignPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "No-sign.png");
            string noSignBase64 = "";
            if (File.Exists(noSignPath))
            {
                var bytes = File.ReadAllBytes(noSignPath);
                noSignBase64 = Convert.ToBase64String(bytes);
            }

            // Filter approvers by type
            var approversLeft = detail.approvelist.Where(a => a.APPROVAL_TYPE_CODE == "ATC02").ToList();
            var approversRight = detail.approvelist.Where(a => a.APPROVAL_TYPE_CODE == "ATC03").ToList();
            int maxRows = Math.Max(approversLeft.Count, approversRight.Count);

            htmlSign.Append(@"
    <div style='width:100%; display: flex; justify-content: center;'>
        <table class='signature-table t-12' style='width:500px; border:none;'>
");

            for (int i = 0; i < maxRows; i++)
            {
                var approver1 = approversLeft.ElementAtOrDefault(i);
                var approver2 = approversRight.ElementAtOrDefault(i);

                string signatureHtml1 = "", signatureHtml2 = "";
                string base64_1 = null, base64_2 = null;

                // Approver 1 signature (ATC02)
                if (approver1 != null)
                {
                    if (flagSign == "Y" && !string.IsNullOrEmpty(approver1.E_Signature) && approver1.E_Signature.Contains("<content>"))
                    {
                        try
                        {
                            var contentStart = approver1.E_Signature.IndexOf("<content>") + "<content>".Length;
                            var contentEnd = approver1.E_Signature.IndexOf("</content>");
                            base64_1 = approver1.E_Signature.Substring(contentStart, contentEnd - contentStart);
                            signatureHtml1 = $@"<div class='t-12 text-center tab1'>
    <img src='data:image/png;base64,{base64_1}' alt='signature' style='max-height: 80px;' />
</div>";
                        }
                        catch
                        {
                            signatureHtml1 = !string.IsNullOrEmpty(noSignBase64)
                                ? $@"<div class='t-12 text-center tab1'>
    <img src='data:image/png;base64,{noSignBase64}' alt='no-signature' style='max-height: 80px;' />
</div>"
                                : "<div class='t-12 text-center tab1'>(ลงชื่อ....................)</div>";
                        }
                    }
                    else
                    {
                        signatureHtml1 = !string.IsNullOrEmpty(noSignBase64)
                            ? $@"<div class='t-12 text-center tab1'>
    <img src='data:image/png;base64,{noSignBase64}' alt='no-signature' style='max-height: 80px;' />
</div>"
                            : "<div class='t-12 text-center tab1'>(ลงชื่อ....................)</div>";
                    }
                }

                // Approver 2 signature (ATC03)
                if (approver2 != null)
                {
                    if (flagSign == "Y" && !string.IsNullOrEmpty(approver2.E_Signature) && approver2.E_Signature.Contains("<content>"))
                    {
                        try
                        {
                            var contentStart = approver2.E_Signature.IndexOf("<content>") + "<content>".Length;
                            var contentEnd = approver2.E_Signature.IndexOf("</content>");
                            base64_2 = approver2.E_Signature.Substring(contentStart, contentEnd - contentStart);
                            signatureHtml2 = $@"<div class='t-12 text-center tab1'>
    <img src='data:image/png;base64,{base64_2}' alt='signature' style='max-height: 80px;' />
</div>";
                        }
                        catch
                        {
                            signatureHtml2 = !string.IsNullOrEmpty(noSignBase64)
                                ? $@"<div class='t-12 text-center tab1'>
    <img src='data:image/png;base64,{noSignBase64}' alt='no-signature' style='max-height: 80px;' />
</div>"
                                : "<div class='t-12 text-center tab1'>(ลงชื่อ....................)</div>";
                        }
                    }
                    else
                    {
                        signatureHtml2 = !string.IsNullOrEmpty(noSignBase64)
                            ? $@"<div class='t-12 text-center tab1'>
    <img src='data:image/png;base64,{noSignBase64}' alt='no-signature' style='max-height: 80px;' />
</div>"
                            : "<div class='t-12 text-center tab1'>(ลงชื่อ....................)</div>";
                    }
                }

                htmlSign.Append($@"
        <tr>
            <td>
                {(approver1 != null ? $@"{signatureHtml1}
                <div>({System.Net.WebUtility.HtmlEncode(approver1.EMPLOYEE_Name ?? "(ชื่อผู้ลงนาม)")})</div>
                <div>{System.Net.WebUtility.HtmlEncode(approver1.EMPLOYEE_PositionName ?? "ตำแหน่ง")}</div>" : "")}
            </td>
            <td>
                {(approver2 != null ? $@"{signatureHtml2}
                <div>({System.Net.WebUtility.HtmlEncode(approver2.EMPLOYEE_Name ?? "(ชื่อผู้ลงนาม)")})</div>
                <div>{System.Net.WebUtility.HtmlEncode(approver2.EMPLOYEE_PositionName ?? "ตำแหน่ง")}</div>" : "")}
            </td>
        </tr>
    ");
            }
            htmlSign.Append(@"
        </table>
    </div>
");
        }


        #endregion Signature Table

        htmlBody.Append($@"
    <div class='text-center t-14'>
        <b>การทบทวนกระบวนการของ {detail.BusinessUnitOwner} ประจำปี {detail.FiscalYear}</b>
    </div>
    <div class='t-12'>ความเป็นมา</div>
</br>
    <div>
        {(string.IsNullOrEmpty(detail.PROCESS_BACKGROUND)
                ? ""
                : string.Join("", detail.PROCESS_BACKGROUND
                    .Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)
                    .Select(line => $"<div class='tab1 t-12'>{System.Net.WebUtility.HtmlEncode(line)}</div>")))}
    </div>
{htmlDescript}
   
    <div class='section-divider'></div>
    <div class='t-12'>การทบทวนกระบวนการของ {detail.BusinessUnitOwner} ประจำปี {detail.FiscalYear} ดังนี้</div>
    <div class='table-container'>
      {htmlTable}
    </div>
    <div class='note t-12'>หมายเหตุ: *ทบทวนตาม JD/ **ทบทวนตาม วค.2/***ทบทวนตามภารกิจงานปัจจุบัน</div>
    {(detail.WorkflowProcesses?.Count > 0
            ? $@"<div class='t-12'>กระบวนการที่จัดทำ Workflow เพิ่มเติม ได้แก่</div>
            <div class='workflow-list'>{string.Join("", detail.WorkflowProcesses.Select(wf => $"<div class='t-12'>• {System.Net.WebUtility.HtmlEncode(wf)}</div>"))}</div>"
            : "")}
    {htmlComment}
</br>   

{htmlSign}

    ");

        

            var html = $@"
<!DOCTYPE html>
<html lang=th>
<head>
    <meta charset='utf-8'>
<style>
    @font-face {{
        font-family: 'TH Sarabun New';
        src: url('data:font/truetype;charset=utf-8;base64,{fontBase64}') format('truetype');
        font-weight: normal;
        font-style: normal;
    }}
    body {{
        font-size: 16px;
        font-family: 'TH Sarabun New', Arial, sans-serif;
        margin: 0;
        padding: 24px;
    }}
    body, p, div {{
        overflow-wrap: break-word;
        -webkit-line-break: after-white-space;
        hyphens: none;
    }}
    .t-12 {{ font-size: 1em; }}
    .t-13 {{ font-size: 1.2em; }}
    .t-14 {{ font-size: 1.3em; }}
    .t-16 {{ font-size: 1.5em; }}
    .t-18 {{ font-size: 1.7em; }}
    .t-20 {{ font-size: 1.9em; }}
    .t-22 {{ font-size: 2.1em; }}
    .section-title {{
        font-size: 1.2em;
        font-weight: bold;
        margin-top: 24px;
        margin-bottom: 8px;
        color: #0056b3;
    }}
    .text-center {{
        text-align: center;
        width: 100%;
        margin-bottom: 24px;
    }}
.text-right {{
    float: right;
    text-align: right;
}}
    .table-container {{
        margin: 24px 0;
    }}
    table {{
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        overflow: hidden;
    }}
    th, td {{
        padding: 10px 8px;
        border: 1px solid #dee2e6;
        word-break: break-word;
        overflow-wrap: anywhere;
        white-space: normal;
        vertical-align: top;
    }}
    .signature-table td {{
        padding: 16px;
        font-size: 1em;
        text-align: center;
        border: none !important;
    }}
    .signature-table {{
    border-radius: 0 !important;
    box-shadow: none !important;
    background: none !important;
}}
    .note {{
        font-style: italic;
        margin-bottom: 12px;
        color: #888;
    }}
    .tab1 {{ text-indent: 48px; }}
    .tab2 {{ text-indent: 96px; }}
    .comment-section {{
        border-radius: 6px;
        padding: 12px 18px;
        margin: 12px 0;
    }}
    .workflow-list {{
        margin-left: 32px;
    }}
    ol {{
        margin-left: 32px;
    }}
    .section-divider {{
        border-bottom: 2px solid #e3e3e3;
        margin: 24px 0 16px 0;
    }}
.signature-container {{display: flex;
        justify-content: flex-end;
        width: 100%;
    }}
</style>
</head>
<body>
    {htmlBody}
</body>
</html>
";
       // _logger.LogInformation(html);
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
        <div class='t-14 text-center'>
           <!-- <b>การทบทวนกลุ่มกระบวนการหลักและกลุ่มกระบวนการสนับสนุน {detail.BusinessUnitOwner} ประจำปี {detail.FiscalYear}</b> -->
                 <b>แผนภาพระบบงาน(Work System) ประจำปี {detail.FiscalYear}</b>

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
            htmlBody.Append("<td rowspan='2' class='t-16' style='width:25%;font-weight:bold;background:#fff;'>กลุ่มกระบวน<br/>การหลัก<br/>(Core Process)</td>");
            foreach (var core in detail.CoreProcesses)
                htmlBody.Append($"<td class='t-16' style='background:#00C896;text-align:center;vertical-align:top;'>{System.Net.WebUtility.HtmlEncode(core.ProcessGroupCode)}</td>");
            htmlBody.Append("</tr>");
            // Row 2: ชื่อกระบวนการ
            htmlBody.Append("<tr>");
            foreach (var core in detail.CoreProcesses)
                htmlBody.Append($"<td class='t-16' style='background:#00C896;text-align:center;vertical-align:top;white-space:normal;word-break:break-word;'>{System.Net.WebUtility.HtmlEncode(core.ProcessGroupName)}</td>");
            htmlBody.Append("</tr>");
            htmlBody.Append("</table>");
        }

        // Supporting Process Table
        if (detail.SupportProcesses != null && detail.SupportProcesses.Count > 0)
        {
            htmlBody.Append("<table class='w-100' border='1' cellpadding='6' style='border-collapse:collapse;margin-bottom:12px; table-layout:fixed;'>");
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
                    htmlBody.Append($"<td class='t-16' rowspan='{detail.SupportProcesses.Count}' style='width:25%;font-weight:bold;'>กลุ่มกระบวนการ<br/>สนับสนุน<br/>(Supporting Process)</td>");
                }
                htmlBody.Append($"<td class='t-16' style='background:#4CB1F0;text-align:center;width:10%;'>{System.Net.WebUtility.HtmlEncode(support.ProcessGroupCode)}</td>");
                htmlBody.Append($"<td class='t-16' style='background:#4CB1F0;text-align:left;width:70%;'>{System.Net.WebUtility.HtmlEncode(support.ProcessGroupName)}</td>");
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
            font-size: 16px;
            font-family: 'THSarabunNew', Arial, sans-serif;
            margin: 0;
            padding: 24px;
        }}
        body, p, div, th, td {{
            word-break: keep-all;
            overflow-wrap: break-word;
            -webkit-line-break: after-white-space;
            hyphens: none;
        }}
        .t-14 {{ font-size: 1.3em; }}
        .t-16 {{ font-size: 1.5em; }}
        .t-18 {{ font-size: 1.7em; }}
        .t-20 {{ font-size: 1.9em; }}
        .t-22 {{ font-size: 2.1em; }}
        .section-title {{
            font-size: 1.2em;
            font-weight: bold;
            margin-top: 24px;
            margin-bottom: 8px;
            color: #0056b3;
        }}
        .text-center {{
            text-align: center;
            width: 100%;
            margin-bottom: 24px;
        }}
        .table-container {{
            margin: 24px 0;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
            overflow: hidden;
        }}
        th, td {{
            padding: 10px 8px;
            border: 1px solid #dee2e6;
            word-break: break-word;
            vertical-align: top;
        }}
        .signature-table td {{
            padding: 16px;
            font-size: 1em;
            text-align: center;
            border: none;
        }}
        .signature-table {{
    border-radius: 0 !important;
    box-shadow: none !important;
    background: none !important;
}}
        .note {{
            font-style: italic;
            margin-bottom: 12px;
            color: #888;
        }}
        .tab1 {{ text-indent: 48px; }}
        .tab2 {{ text-indent: 96px; }}
        .comment-section {{
            border-radius: 6px;
            padding: 12px 18px;
            margin: 12px 0;
        }}
        .workflow-list {{
            margin-left: 32px;
        }}
        ol {{
            margin-left: 32px;
        }}
        .section-divider {{
            border-bottom: 2px solid #e3e3e3;
            margin: 24px 0 16px 0;
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
        <div class='t-16 text-center'>
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
            htmlBody.Append("<td rowspan='2' class='t-16' style='width:25%;font-weight:bold;background:#fff;'>กลุ่มกระบวน<br/>การหลัก<br/>(Core Process)</td>");
            foreach (var core in detail.CoreProcesses)
                htmlBody.Append($"<td class='t-16' style='background:#00C896;text-align:center;vertical-align:middle;'>{System.Net.WebUtility.HtmlEncode(core.ProcessGroupCode)}</td>");
            htmlBody.Append("</tr>");
            // Row 2: ชื่อกระบวนการ
            htmlBody.Append("<tr>");
            foreach (var core in detail.CoreProcesses)
                htmlBody.Append($"<td class='t-16' style='background:#00C896;text-align:center;vertical-align:middle;white-space:normal;word-break:break-word;'>{System.Net.WebUtility.HtmlEncode(core.ProcessGroupName)}</td>");
            htmlBody.Append("</tr>");
            htmlBody.Append("</table>");
        }

        // Supporting Process Table
        if (detail.SupportProcesses != null && detail.SupportProcesses.Count > 0)
        {
            htmlBody.Append("<table class='w-100' border='1' cellpadding='6' style='border-collapse:collapse;margin-bottom:12px; table-layout:fixed;'>");
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
                    htmlBody.Append($"<td class='t-16' rowspan='{detail.SupportProcesses.Count}' style='width:25%;font-weight:bold;'>กลุ่มกระบวนการ<br/>สนับสนุน<br/>(Supporting Process)</td>");
                }
                htmlBody.Append($"<td class='t-16' style='background:#4CB1F0;text-align:center;width:10%;'>{System.Net.WebUtility.HtmlEncode(support.ProcessGroupCode)}</td>");
                htmlBody.Append($"<td class='t-16' style='background:#4CB1F0;text-align:left;width:70%;'>{System.Net.WebUtility.HtmlEncode(support.ProcessGroupName)}</td>");
                htmlBody.Append("</tr>");
            }
            htmlBody.Append("</table>");
        }

        // Section heading
        htmlBody.Append($@"
        <div>
            การทบทวนกระบวนการของ {detail.BusinessUnitOwner} ประจำปี {detail.FiscalYear} ดังนี้
        </div>
    ");

        // Three-column table
        int rowCount = Math.Max(
            Math.Max(detail.PrevProcesses?.Count ?? 0, detail.CurrentProcesses?.Count ?? 0),
            detail.ControlActivities?.Count ?? 0
        );
        htmlBody.Append("<table class='w-100' border='1' cellpadding='6' style='border-collapse:collapse;margin-bottom:12px;'>");
        htmlBody.Append($@"
        <tr style='background:#DDEBF7;font-weight:bold;text-align:center;'>
            <td>กระบวนการ ปี {detail.FiscalYearPrevious} (เดิม)</td>
            <td>กระบวนการ ปี {detail.FiscalYear} (ทบทวน)</td>
            <td>กิจกรรมควบคุม (Control Activity)</td>
        </tr>
    ");
        for (int i = 0; i < rowCount; i++)
        {
            htmlBody.Append("<tr>");
            htmlBody.Append($"<td>{System.Net.WebUtility.HtmlEncode(detail.PrevProcesses?.ElementAtOrDefault(i) ?? "")}</td>");
            htmlBody.Append($"<td>{System.Net.WebUtility.HtmlEncode(detail.CurrentProcesses?.ElementAtOrDefault(i) ?? "")}</td>");
            htmlBody.Append($"<td>{System.Net.WebUtility.HtmlEncode(detail.ControlActivities?.ElementAtOrDefault(i) ?? "")}</td>");
            htmlBody.Append("</tr>");
        }
        htmlBody.Append("</table>");

        // Note
        htmlBody.Append("<div style='font-style:italic;margin-bottom:12px;'>หมายเหตุ: *ทบทวนตาม JD/ **ทบทวนตาม คว.2/***ทบทวนตามภารกิจงานปัจจุบัน</div>");

        // Workflow processes
        if (detail.WorkflowProcesses?.Count > 0)
        {
            htmlBody.Append("<div>กระบวนการที่จัดทำ Workflow เพิ่มเติม ได้แก่</div>");
            foreach (var wf in detail.WorkflowProcesses)
                htmlBody.Append($"<div style='margin-left:32px;'>• {System.Net.WebUtility.HtmlEncode(wf)}</div>");
        }

        // Comments
        htmlBody.Append("<div>ความคิดเห็น</div>");
        htmlBody.Append("<div style='margin-left:32px;'>☐ เห็นชอบการปรับปรุง</div>");
        htmlBody.Append("<div style='margin-left:32px;'>☐ มีความเห็นเพิ่มเติม</div>");

        // Approve remarks
        if (detail.ApproveRemarks?.Length > 0)
        {
            foreach (var r in detail.ApproveRemarks)
                htmlBody.Append($"<div style='margin-left:32px;'>{System.Net.WebUtility.HtmlEncode(r)}</div>");
        }

        // Signature section
        htmlBody.Append(@"
        <table class='signature-table'>
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
            .t-16 {{ font-size: 1.5em; }}
            .t-18 {{ font-size: 1.7em; }}
            .t-22 {{ font-size: 1.9em; }}
            .tab1 {{ text-indent: 48px; }}
            .tab2 {{ text-indent: 96px; }}
            .tab3 {{ text-indent: 144px; }}
            .tab4 {{ text-indent: 192px; }}
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
}
