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
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");
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
        var signatoryHtml = new StringBuilder();
        var companySealHtml = new StringBuilder();
        bool sealAdded = false; // กันซ้ำ

        var dataSignatories = signlist.Where(e => e?.Signatory_Type != null).ToList();
        // Group signatories
        var dataSignatoriesTypeOSMEP = dataSignatories
            .Where(e => e.Signatory_Type == "OSMEP_S" || e.Signatory_Type == "OSMEP_W")
            .ToList();
        var dataSignatoriesTypeCP = dataSignatories
            .Where(e => e.Signatory_Type == "CP_S" || e.Signatory_Type == "CP_W")
            .ToList();

        // Helper to render a signatory block
        string RenderSignatory(E_ConReport_SignatoryModels signer)
        {
            string signatureHtml;
            string noSignPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "No-sign.png");
            string noSignBase64 = "";
            if (File.Exists(noSignPath))
            {
                var bytes = File.ReadAllBytes(noSignPath);
                noSignBase64 = Convert.ToBase64String(bytes);
            }

            if (!string.IsNullOrEmpty(signer?.DS_FILE) && signer.DS_FILE.Contains("<content>"))
            {
                try
                {
                    var contentStart = signer.DS_FILE.IndexOf("<content>") + "<content>".Length;
                    var contentEnd = signer.DS_FILE.IndexOf("</content>");
                    var base64 = signer.DS_FILE.Substring(contentStart, contentEnd - contentStart);

                    signatureHtml = $@"<div class='t-16 text-center tab1'>
    <img src='data:image/png;base64,{base64}' alt='signature' style='max-height: 80px;' />
</div>";
                }
                catch
                {
                    signatureHtml = !string.IsNullOrEmpty(noSignBase64)
                        ? $@"<div class='t-16 text-center tab1'>
    <img src='data:image/png;base64,{noSignBase64}' alt='no-signature' style='max-height: 80px;' />
</div>"
                        : "<div class='t-16 text-center tab1'>(ลงชื่อ....................)</div>";
                }
            }
            else
            {
                signatureHtml = !string.IsNullOrEmpty(noSignBase64)
                    ? $@"<div class='t-16 text-center tab1'>
    <img src='data:image/png;base64,{noSignBase64}' alt='no-signature' style='max-height: 80px;' />
</div>"
                    : "<div class='t-16 text-center tab1'>(ลงชื่อ....................)</div>";
            }

            string name = signer?.Signatory_Name ?? "";
            string nameBlock = (signer?.Signatory_Type != null && signer.Signatory_Type.EndsWith("_W"))
                ? $"({name})พยาน"
                : $"({name})";

            return $@"
<div class='sign-single-right'>
    {signatureHtml}
    <div class='t-16 text-center tab1'>{nameBlock}</div>
    <div class='t-16 text-center tab1'>{signer?.Position}</div>
</div>";
        }

        // Build HTML for each column
        var smeSignHtml = new StringBuilder();
        foreach (var signer in dataSignatoriesTypeOSMEP)
        {
            smeSignHtml.AppendLine(RenderSignatory(signer));
        }

        var customerSignHtml = new StringBuilder();
        foreach (var signer in dataSignatoriesTypeCP)
        {
            customerSignHtml.AppendLine(RenderSignatory(signer));
        }
        //คราประทับ
        var companySealSignatory = dataSignatoriesTypeCP.Where(e => e.Company_Seal != null).FirstOrDefault();
        if (companySealSignatory != null && !string.IsNullOrEmpty(companySealSignatory.Company_Seal) && companySealSignatory.Company_Seal.Contains("<content>"))
        {
            try
            {
                var contentStart = companySealSignatory.Company_Seal.IndexOf("<content>") + "<content>".Length;
                var contentEnd = companySealSignatory.Company_Seal.IndexOf("</content>");
                var base64 = companySealSignatory.Company_Seal.Substring(contentStart, contentEnd - contentStart);

                var companySeal = $@"
<div class='t-16 text-center tab1'>
    <img src='data:image/png;base64,{base64}' alt='signature' style='max-height: 80px;' />
</div>";

                companySealHtml.AppendLine($@"
<div class='text-center'>
    {companySeal}
</div>
");
                sealAdded = true;
            }
            catch
            {
                companySealHtml.AppendLine("<div class='t-16 text-center tab1'></div>");
                sealAdded = true;
            }
        }
        else
        {
            // ไม่มีไฟล์ตรา/ไม่มี <content> ⇒ ใส่ placeholder ครั้งเดียว
            companySealHtml.AppendLine("<div class='t-16 text-center tab1'></div>");
            sealAdded = true;
        }

        // Output as a table
        var signatoryTableHtml = $@"
<table class='signature-table'>
    <tr>
        <td style='width:50%; vertical-align:top;'>
            
            {smeSignHtml}
        </td>
        <td style='width:50%; vertical-align:top;'>
           
            {customerSignHtml}
     {companySealHtml}
        </td>
    </tr>
</table>

";
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
            if (!existingClass.Contains("t-16"))
            {
                p.SetAttributeValue("class", (existingClass + " t-16").Trim());
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
        src: url('file:///{fontPath}') format('truetype');
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
    .t-16 {{ font-size: 1.5em; !important;
line-height: 1.6; !important;
}}
    .t-18 {{ font-size: 1.7em; }}
    .t-22 {{ font-size: 1.9em; }}
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
 font: inherit !important;
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
</br>

</br>
    <div class='t-22 text-center'><b>แนวทางการจัดทำ</b></div>
    <div class='t-22 text-center'><b>เอกสารแนบท้ายบันทึกข้อตกลงความร่วมมือและสัญญาร่วมดำเนินการ</b></div>
    <div class='t-18 text-center'><b>ข้อกำหนดของการดำเนินงาน</b></div>
  <div class='t-18 text-center'><b>โครงการ {dataResult.Contract_Name}</b></div>
  <div class='t-18 text-center'><b>ระหว่าง</b></div>
   <div class='t-18 text-center'><b>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ( สสว. )</b></div>
 <div class='t-18 text-center'><b>กับ</b></div>
<div class='t-18 text-center'><b>ชื่อหน่วยร่วมดำเนินการ {dataResult.Start_Unit} </b></div>
</br>

<p class='t-16 tab0'><b>๑.รายละเอียด</b></p>
<div class='t-16 editor-content'>
    {cleanedHtml}
</div>
    
 


<P class='t-16 tab0'><b>๒. เงื่อนไขอื่น ๆ</b></P>

<P class='t-16 tab1'>๒.๑ (ชื่อย่อหน่วยร่วมดำเนินการ) ต้องปฏิบัติตามคู่มือการดำเนินโครงการ ของ สสว. โดยเคร่งครัด</P>
<P class='t-16 tab1'>๒.๒ (ชื่อย่อหน่วยร่วมดำเนินการ) ต้องประสานงานกับ สสว. อย่างต่อเนื่องและใกล้ชิด  ต้องอำนวยความสะดวกให้ สสว. หรือเจ้าหน้าที่ของ สสว. ในการประสาน กำกับ บริหารจัดการ และประเมินผลการดำเนินโครงการ</P>
<P class='t-16 tab1'>๒.๓ (ชื่อย่อหน่วยร่วมดำเนินการ) ต้องรับผิดชอบประสานงานกับผู้ประกอบการที่เข้าร่วมโครงการ อย่างต่อเนื่องใกล้ชิด  ต้องอำนวยความสะดวกให้ผู้ประกอบการที่เข้าร่วมโครงการและการดำเนินกิจกรรมตามโครงการ  รวมถึงสนับสนุนค่าใช้จ่ายในส่วนที่เกี่ยวข้องกับกิจกรรมในโครงการให้แก่ผู้ประกอบการตามที่ สสว. กำหนดไว้ (ถ้ามี)</P>
<P class='t-16 tab1'>๒.๔ กรณีค่าใช้จ่ายต่าง ๆ ในการดำเนินการและอื่น ๆ อันเกิดขึ้นจากการดำเนินกิจกรรมกับ สสว. (ชื่อย่อหน่วยร่วมดำเนินการ) ต้องเป็นผู้รับผิดชอบค่าใช้จ่ายทั้งสิ้น จะใช้สิทธิเรียกร้องค่าเสียหายใด ๆ จาก สสว. ไม่ได้</P>
<P class='t-16 tab1'>๒.๕ (ชื่อย่อหน่วยร่วมดำเนินการ) มีหน้าที่รับผิดชอบบริหารจัดการบัญชีและการจัดเก็บเอกสารที่เกี่ยวข้องกับการดำเนินโครงการ และเกี่ยวกับการรับเงิน การจ่ายเงินหรือการก่อหนี้ผูกพันทางการเงิน จัดทำบัญชีรายรับรายจ่าย รวมถึงหนังสือและเอกสารอื่นที่เกี่ยวข้องกับการดำเนินโครงการ เพื่อให้หน่วยงานตรวจสอบ เช่น สำนักงานการตรวจเงินแผ่นดิน สำนักงานป้องกันและปราบปรามการทุจริตแห่งชาติ เป็นต้น สามารถใช้ตรวจสอบและอ้างอิงได้  ทั้งนี้ ระยะเวลาการจัดเก็บเอกสารให้เป็นไปตามระเบียบของราชการ</P>

<P class='t-16 tab0'><b>๓. เอกสารประกอบการจัดทำบันทึกข้อตกลงความร่วมมือและสัญญาร่วมดำเนินการ</b></P>
<P class='t-16 tab1'>สสว. (รับรองสำเนาถูกต้อง ทุกสำเนาเอกสาร)</P>
<P class='t-16 tab2'>๑. สำเนาหนังสือแต่งตั้งผู้มีอำนาจลงนามของ สสว. จำนวน ๒ ชุด </P>
<P class='t-16 tab2'>๒. สำเนาบัตรข้าราชการหรือบัตรประชาชนของผู้มีอำนาจลงนามของ สสว. จำนวน ๒ ชุด</P>
<P class='t-16 tab2'>๓. กรณีมอบอำนาจแทนผู้มีอำนาจลงนาม ต้องมีหนังสือมอบอำนาจ และสำเนาบัตรประชาชน
ผู้มีอำนาจลงนามและผู้รับมอบอำนาจลงนาม เพิ่มจำนวน ๑ ชุด</P>
<P class='t-16 tab1'>หน่วยร่วมดำเนินการ (รับรองสำเนาถูกต้อง ทุกสำเนาเอกสาร)</P>
<P class='t-16 tab2'>๑. สำเนาเอกสารแสดงการจดทะเบียนเป็นนิติบุคคล 或แสดงการจัดตั้งหน่วยงาน จำนวน ๒ ชุด </P>
<P class='t-16 tab2'>๒. สำเนาหนังสือแต่งตั้งผู้มีอำนาจลงนาม จำนวน ๒ ชุด</P>
<P class='t-16 tab2'>๔. กรณีมอบอำนาจแทนผู้มีอำนาจลงนาม ต้องมีหนังสือมอบอำนาจ และสำเนาบัตรประชาชน
ผู้มีอำนาจลงนามและผู้รับมอบอำนาจลงนาม เพิ่มจำนวน ๑ ชุด
</P>

</div>
{signatoryTableHtml}
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
