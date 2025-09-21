using BatchAndReport.DAO;
using BatchAndReport.Models;
using DinkToPdf.Contracts;
using System.Text;
using System.Text.RegularExpressions;
public class WordEContract_AMJOAService
{
    private readonly WordServiceSetting _w;
    private readonly E_ContractReportDAO _eContractReportDAO;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    private readonly Econtract_Report_AMJOADAO _eContractReportAMJOADAO;

    public WordEContract_AMJOAService(
        WordServiceSetting ws,
        E_ContractReportDAO eContractReportDAO
      , IConverter pdfConverter
        , Econtract_Report_AMJOADAO eContractReportAMJOADAO
    )
    {
        _w = ws;
        _eContractReportDAO = eContractReportDAO;
        _pdfConverter = pdfConverter;
        _eContractReportAMJOADAO = eContractReportAMJOADAO;
    }


    public async Task<string> OnGetWordContact_AMJOAServiceHtmlToPDF(string conId)
    {
        var dataResult = await _eContractReportAMJOADAO.GetAMJOAAsync(conId);
        if (dataResult == null)
            throw new Exception("AMJOA data not found.");
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf");
        string fontBase64 = "";
        if (File.Exists(fontPath))
        {
            var bytes = File.ReadAllBytes(fontPath);
            fontBase64 = Convert.ToBase64String(bytes);
        }
        var cssPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "css", "contract.css").Replace("\\", "/");
        var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
        string logoBase64 = "";
        if (System.IO.File.Exists(logoPath))
        {
            var bytes = System.IO.File.ReadAllBytes(logoPath);
            logoBase64 = Convert.ToBase64String(bytes);
        }
        string contractLogoHtml;
        if (!string.IsNullOrEmpty(dataResult.Organization_Logo) && dataResult.Organization_Logo.Contains("<content>"))
        {
            try
            {
                // ตัดเอาเฉพาะ Base64 ในแท็ก <content>...</content>
                var contentStart = dataResult.Organization_Logo.IndexOf("<content>") + "<content>".Length;
                var contentEnd = dataResult.Organization_Logo.IndexOf("</content>");
                var contractlogoBase64 = dataResult.Organization_Logo.Substring(contentStart, contentEnd - contentStart);

                contractLogoHtml = $@"<div style='display:inline-block; padding:20px; font-size:32pt;'>
             <img src='data:image/jpeg;base64,{contractlogoBase64}' width='240' height='80' />
            </div>";
            }
            catch
            {
                contractLogoHtml = "";
            }
        }
        else
        {
            contractLogoHtml = "";
        }

      


        #region signlist 

        var signlist = await _eContractReportDAO.GetSignNameAsync(conId, "AMJOA");
        // call function RenderSignatory
        var signatoryTableHtml = "";
        if (signlist.Count > 0)
        {
            signatoryTableHtml = await _eContractReportDAO.RenderSignatory(signlist);

        }
        var signatoryTableHtmlWitnesses = "";

        if (signlist.Count > 0)
        {
            signatoryTableHtmlWitnesses = await _eContractReportDAO.RenderSignatory_Witnesses(signlist);
        }
        #endregion signlist

        #region
        // ตัวอย่างการใช้ Regex เพื่อลบ style attribute ออก
        var cleanDescription = Regex.Replace(dataResult.Contract_Description, "style=\"[^\"]*\"", string.Empty);

        // หรือใช้ HtmlAgilityPack ที่แนะนำมากกว่า
        var htmlDoc = new HtmlAgilityPack.HtmlDocument();
        htmlDoc.LoadHtml(cleanDescription);

        // ลบ style และ dir attributes ออกจากทุก element
        foreach (var element in htmlDoc.DocumentNode.DescendantsAndSelf())
        {
            if (element.Attributes.Contains("style"))
            {
                element.Attributes["style"].Remove();
            }
            if (element.Attributes.Contains("dir"))
            {
                element.Attributes["dir"].Remove();
            }
        }

        // ลบ tag <span> ที่ไม่มี attribute หรือ class และคงไว้แต่ข้อความ
        foreach (var span in htmlDoc.DocumentNode.Descendants("span").ToList())
        {
            // ตรวจสอบว่า tag <span> ไม่มี attributes และไม่มี child nodes ที่เป็น element (มีแค่ text)
            if (!span.HasAttributes && !span.ChildNodes.Any(n => n.NodeType == HtmlAgilityPack.HtmlNodeType.Element))
            {
                var textNode = htmlDoc.CreateTextNode(span.InnerText);
                span.ParentNode.ReplaceChild(textNode, span);
            }
        }
        // ⭐ เพิ่ม class="t-16" ให้กับทุก <p>
        foreach (var p in htmlDoc.DocumentNode.Descendants("p"))
        {
            var existingClass = p.GetAttributeValue("class", "");
            if (!existingClass.Contains("t-14"))
            {
                p.SetAttributeValue("class", (existingClass + " t-14").Trim());
            }
        }

        string cleanedHtml = htmlDoc.DocumentNode.OuterHtml;

        #endregion

        var html = $@"<html>
<head>
    <meta charset='utf-8'>
  
    <style>
    @font-face {{
        font-family: 'TH Sarabun New';
                src: url('data:font/truetype;charset=utf-8;base64,{fontBase64}') format('truetype');
        font-weight: normal;
        font-style: normal;
    }}
  body {{font-family: 'TH Sarabun New', Arial, Tahoma, sans-serif !important;
    font-size: 22px !important;
    color: #000 !important;
    word-break: keep-all;
    overflow-wrap: break-word;
    -webkit-line-break: after-white-space;
    hyphens: none;
}}
    body, p, div {{
    font-family: 'TH Sarabun New', Arial, Tahoma, sans-serif !important;
    font-size: 22px !important;
    color: #000 !important;
    word-break: keep-all;
    overflow-wrap: break-word;
    -webkit-line-break: after-white-space;
    hyphens: none;
    }}
 .t-12 {{ font-size: 1em; }}
        .t-14 {{ font-size: 1.1em; }}

    .t-16 {{ font-size: 1.5em; }}


    .t-18 {{ font-size: 1.7em; !important; }}
    .t-22 {{ font-size: 1.9em;!important; }}
    .tab0 {{ text-indent: 0px; }}
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
    .table {{ width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 28pt; }}
    .table th, .table td {{ border: 1px solid #000; padding: 8px; }}
    .sign-double {{ display: flex; }}
    .text-center-right-brake {{ margin-left: 50%; }}
    .text-right {{ text-align: right; }}
    .contract, .section {{ margin: 12px 0; line-height: 1.7; }}
    .section {{ font-weight: bold; font-size: 1.2em; text-align: left; margin-top: 24px; }}
    .signature-table {{ width: 100%; margin-top: 32px; border-collapse: collapse; }}
    .signature-table td {{ padding: 16px; text-align: center; vertical-align: top; font-size: 1.1em; }}
    .logo-table {{ width: 100%; border-collapse: collapse; margin-top: 40px; }}
    .logo-table td {{ border: none; }}
    p {{ margin: 0; padding: 0; }}
.editor-content,
.editor-content p,
.editor-content span,
.editor-content li {{

    color: inherit !important;
}}
    body, p, div, span, li, td, th, table, b, strong, h1, h2, h3, h4, h5, h6 {{
        font-family: 'TH Sarabun New', Arial, Tahoma, sans-serif !important;
      
        color: #000 !important;
    }}
</style>
</head><body>

<table style='width:100%; border-collapse:collapse; margin-top:40px;'>
    <tr>
        <!-- Left: SME logo -->
        <td style='width:60%; text-align:left; vertical-align:top;'>
        <div style='display:inline-block;  padding:20px; font-size:32pt;'>
             <img src='data:image/jpeg;base64,{logoBase64}' width='240' height='80' />
            </div>
        </td>
        <!-- Right: Contract code box (replace with your actual contract code if needed) -->
        <td style='width:40%; text-align:center; vertical-align:top;'>
          
        </td>
    </tr>
</table>

    <h2  class='t-16 text-center'><b>{dataResult.Contract_Title}</b></h2 >
  <h2  class='t-16 text-center'><b>โครงการ {dataResult.Contract_Name}</b></h2 >
  <h2  class='t-16 text-center'><b>ระหว่าง</b></div>
   <h2  class='t-16 text-center'><b>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ( สสว. )</b></h2 >
 <h2  class='t-16 text-center'><b>กับ</b></div>
<h2  class='t-16 text-center'><b>ชื่อหน่วยร่วมดำเนินการ {dataResult.Start_Unit} </b></h2 >
</br>
<div class='t-14 editor-content'>
    {cleanedHtml}
</div>

</div>

<P class='t-14 tab3'>บันทึกข้อตกลงนี้ทำขึ้นเป็นบันทึกข้อตกลงอิเล็กทรอนิกส์ คู่ตกลงได้อ่าน เข้าใจเงื่อนไข และยอมรับเงื่อนไข และได้ยืนยันว่าเป็นผู้มีอำนาจลงนามในบันทึกข้อตกลง จึงได้ลงลายมืออิเล็กทรอนิกส์พร้อมทั้งประทับตรา (ถ้ามี) ในสัญญาไว้ และต่างฝ่ายต่างยึดถือไว้ฝ่ายละหนึ่งฉบับในระบบของตน</P>
</br>
</br>
{signatoryTableHtml}
   <P class='t-14 tab3'>ข้าพเจ้าขอรับรองว่า ทั้งสองฝ่ายได้ลงนามในบันทึกข้อตกลงโดยวิธีการอิเล็กทรอนิกส์ เพื่อแสดงเจตนาของคู่ตกลงแล้ว ข้าพเจ้าจึงได้ลงลายมือชื่ออิเล็กทรอนิกส์รับรองเป็นพยานในบันทึกข้อตกลงพร้อมนี้</P>

{signatoryTableHtmlWitnesses}
</body>
</html>
";


        return html;
    }

    public static ProjectData GenerateMockData()
    {
        var projectData = new ProjectData();

        // Mock data for "Outputs"
        projectData.Outputs = new List<ProjectIndicator>
        {
            new ProjectIndicator
            {
                Description = "เอกสารคู่มือการใช้งาน (User manual documents)",
                UnitOfMeasurement = "ฉบับ (Copies)",
                Target = "5",
                MeasurementMethod = "การตรวจสอบเอกสาร (Document review)"
            },
            new ProjectIndicator
            {
                Description = "จำนวนผู้เข้าอบรม (Number of trainees)",
                UnitOfMeasurement = "คน (People)",
                Target = "100",
                MeasurementMethod = "รายชื่อผู้เข้าอบรม (Trainee list)"
            },
            new ProjectIndicator
            {
                Description = "การจัดกิจกรรม Workshop (Workshop activities)",
                UnitOfMeasurement = "ครั้ง (Times)",
                Target = "2",
                MeasurementMethod = "สรุปผลการจัดกิจกรรม (Activity summary report)"
            },
            new ProjectIndicator
            {
                Description = "รายงานผลการดำเนินงาน (Project progress reports)",
                UnitOfMeasurement = "ฉบับ (Copies)",
                Target = "3",
                MeasurementMethod = "การส่งมอบเอกสารตามกำหนด (On-time document submission)"
            }
        };

        // Mock data for "Outcomes"
        projectData.Outcomes = new List<ProjectIndicator>
        {
            new ProjectIndicator
            {
                Description = "ผู้ใช้งานมีความเข้าใจในการใช้ระบบ (User understanding of the system)",
                UnitOfMeasurement = "เปอร์เซ็นต์ (%)",
                Target = "85%",
                MeasurementMethod = "แบบสอบถามประเมินความพึงพอใจ (Satisfaction survey)"
            },
            new ProjectIndicator
            {
                Description = "ลดระยะเวลาในการทำงาน (Reduced work duration)",
                UnitOfMeasurement = "เปอร์เซ็นต์ (%)",
                Target = "20%",
                MeasurementMethod = "การเปรียบเทียบระยะเวลาก่อนและหลังใช้งาน (Before and after usage time comparison)"
            }
        };

        return projectData;
    }
    public class ProjectIndicator
    {
        public string Description { get; set; }
        public string UnitOfMeasurement { get; set; }
        public string Target { get; set; }
        public string MeasurementMethod { get; set; }
    }

    // Defines a container for the project's success indicators
    public class ProjectData
    {
        public List<ProjectIndicator> Outputs { get; set; }
        public List<ProjectIndicator> Outcomes { get; set; }
    }
}
