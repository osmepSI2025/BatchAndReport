using BatchAndReport.DAO;
using BatchAndReport.Models;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using System.Text;
using System.Threading.Tasks;

public class WordEContract_MemorandumService
{
    private readonly WordServiceSetting _w;
    private readonly E_ContractReportDAO _eContractReportDAO;
    private readonly IConverter _pdfConverter; 
    public WordEContract_MemorandumService(WordServiceSetting ws
            , E_ContractReportDAO eContractReportDAO
        , IConverter pdfConverter
        )
    {
        _w = ws;
        _eContractReportDAO = eContractReportDAO;
        _pdfConverter = pdfConverter;
    }
    #region 4.1.1.2.3.บันทึกข้อตกลงความร่วมมือ
  

    public async Task<string> OnGetWordContact_MemorandumService_HtmlToPDF(string id,string typeContact)
    {
        var result = await _eContractReportDAO.GetMOUAsync(id);

        if (result == null)
        {
            throw new Exception("ไม่พบข้อมูลบันทึกข้อตกลงความร่วมมือ");
        }

        // Logo
        string strContract_Value =  CommonDAO.NumberToThaiText(result.Contract_Value ?? 0);
        string strSign_Date = CommonDAO.ToArabicDateStringCovert(result.Sign_Date ?? DateTime.Now);
        string strStart_Date = CommonDAO.ToArabicDateStringCovert(result.Start_Date ?? DateTime.Now);
        string strEnd_Date = CommonDAO.ToArabicDateStringCovert(result.End_Date ?? DateTime.Now);

        var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
        string logoBase64 = "";
        if (System.IO.File.Exists(logoPath))
        {
            var bytes = System.IO.File.ReadAllBytes(logoPath);
            logoBase64 = Convert.ToBase64String(bytes);
        }
        string contractLogoHtml;
        if (!string.IsNullOrEmpty(result.Organization_Logo) && result.Organization_Logo.Contains("<content>"))
        {
            try
            {
                // ตัดเอาเฉพาะ Base64 ในแท็ก <content>...</content>
                var contentStart = result.Organization_Logo.IndexOf("<content>") + "<content>".Length;
                var contentEnd = result.Organization_Logo.IndexOf("</content>");
                var contractlogoBase64 = result.Organization_Logo.Substring(contentStart, contentEnd - contentStart);

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
        #region checkมอบอำนาจ
        string strAttorneyLetterDate = CommonDAO.ToArabicDateStringCovert(result.Effective_Date ?? DateTime.Now);
        string strAttorneyLetterDate_CP = CommonDAO.ToArabicDateStringCovert(result.CP_S_AttorneyLetterDate ?? DateTime.Now);
        string strAttorneyOsmep = "";
        var HtmlAttorneyOsmep = new StringBuilder();
        if (result.AttorneyFlag == true)
        {
            strAttorneyOsmep = "ผู้มีอำนาจกระทำการแทนปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ เลขคำสั่งสำนักงานที่ " + result.AttorneyLetterNumber + " ฉบับลงวันที่ " + strAttorneyLetterDate + "";

        }
        else
        {
            strAttorneyOsmep = "";
        }
        string strAttorney = "";
        var HtmlAttorney = new StringBuilder();
        if (result.AttorneyFlag == true)
        {
            strAttorney = "ผู้มีอำนาจกระทำการแทนปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ ฉบับลงวันที่ " + strAttorneyLetterDate + "";

        }
        else
        {
            strAttorney = "";
        }
        #endregion

        // Font
       // var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf");
        string fontBase64 = "";
        if (File.Exists(fontPath))
        {
            var bytes = File.ReadAllBytes(fontPath);
            fontBase64 = Convert.ToBase64String(bytes);
        }
        // Purpose list
        var purposeList = await _eContractReportDAO.GetMOUPoposeAsync(id);


        #region signlist 

        var signlist = await _eContractReportDAO.GetSignNameAsync(id, typeContact);
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

        var html = $@"
<html>
<head>
    <meta charset='utf-8'>
       <style>
        @font-face {{
            font-family: 'THSarabunNew';
              src: url('data:font/truetype;charset=utf-8;base64,{fontBase64}') format('truetype');
            font-weight: normal;
            font-style: normal;
        }}
        body {{
            font-size: 22px;
            font-family: 'THSarabunNew', Arial, sans-serif;
        }}
        /* แก้การตัดคำไทย: ไม่หั่นกลางคำ, ตัดเมื่อจำเป็น */
        body, p, div {{
            word-break: keep-all;            /* ห้ามตัดกลางคำ */
            overflow-wrap: break-word;       /* ตัดเฉพาะเมื่อจำเป็น (ยาวจนล้นบรรทัด) */
            -webkit-line-break: after-white-space; /* ช่วย WebKit เก่าจัดบรรทัด */
            hyphens: none;
        }}
         .t-12 {{ font-size: 1em; }}
        .t-14 {{ font-size: 1.1em; }}
        .t-15 {{ font-size: 1.2em; }}
        .t-16 {{ font-size: 1.3em; }}
        .t-18 {{ font-size: 1.7em; }}
        .t-20 {{ font-size: 1.9em; }}
        .t-22 {{ font-size: 1.37em; }}

 .tab0 {{text-align: justify;  }}
           .tab1 {{ text-indent: 48px; text-align: justify;  }}
        .tab2 {{ text-indent: 96px;  text-align: left; }}
        .tab3 {{ text-indent: 144px; text-align: left; }}
        .tab4 {{ text-indent: 192px;  text-align: left;}}
       .normal {{text-align: justify;
        text-align-last: justify;
        width: 100%;
        display: block;
        min-width: 100%;
  letter-spacing: 0.1em; /* เพิ่มช่องไฟเล็กน้อย */
    }}
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
            font-size: 1.4em;
        }}

.logo-table {{ width: 100%; border-collapse: collapse; margin-top: 40px; }}
        .logo-table td {{ border: none; }}
        p {{
            margin: 0;
            padding: 0;
        }}
    </style>
</head>
<body>
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
            {contractLogoHtml}
        </td>
    </tr>
</table>
</br>
</br>
    <div class='t-22 text-center'><B>บันทึกข้อตกลงความร่วมมือ</B></div>
   <div class='t-22 text-center'><B>โครงการ {result.ProjectTitle}</B></div>
    <div class='t-12 text-center'><B>ระหว่าง</B></div>
    <div class='t-12 text-center'><B>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B></div>
    <div class='t-12 text-center'><B>กับ</B></div>
    <div class='t-15 text-center'><B>{result.OrgName ?? ""}</B></div>
    <br/>
     <P class='t-12 tab2'>บันทึกข้อตกลงความร่วมมือฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลาง และขนาดย่อม เมื่อ {strSign_Date} ระหว่าง</P>
    <P class='t-12 tab2'><B>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B>  โดย {result.OSMEP_NAME} ตำแหน่ง {result.OSMEP_POSITION} {strAttorneyOsmep} สำนักงานตั้งอยู่เลขที่ 120 หมู่ 3 ศูนย์ราชการเฉลิมพระเกียรติ 80 พรรษา 5 ธันวาคม 2550 (อาคารซี) ชั้น 2, 10, 11 ถนนแจ้งวัฒนะ แขวงทุ่งสองห้อง เขตหลักสี่ กรุงเทพ 10210 ซึ่งต่อไป ในสัญญาฉบับนี้จะเรียกว่า“สสว.” ฝ่ายหนึ่ง กับ</P>
    <P class='t-12 tab2'>“{result.OrgCommonName ?? ""}” {result.CP_S_NAME} ตำแหน่ง {result.CP_S_POSITION} {strAttorney} สำนักงานตั้งอยู่เลขที่ {result.Office_Loc} ซึ่งต่อไปในสัญญาฉบับนี้จะเรียกว่า “{result.OrgName ?? ""}” อีกฝ่ายหนึ่ง</P>
    <P class='t-12 tab0'><b>วัตถุประสงค์ของความร่วมมือ</b></P>
    <P class='t-12 tab2'>ทั้งสองฝ่ายมีความประสงค์ที่จะร่วมมือกันเพื่อดำเนินการภายใต้โครงการ {result.ProjectTitle} ซึ่งในบันทึกข้อตกลงฉบับนี้ต่อไปจะเรียกว่า “โครงการ” โดยมีรายละเอียดโครงการแผนการดำเนินงาน แผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) และบรรดาเอกสารแนบท้ายบันทึกข้อตกลงฉบับนี้ ซึ่งให้ถือเป็นส่วนหนึ่งของบันทึกข้อตกลงฉบับนี้ มีระยะเวลา ตั้งแต่วันที่ {strStart_Date} จนถึงวันที่ {strEnd_Date} โดยมีวัตถุประสงค์ ในการดำเนินโครงการ ดังนี้</P>
{(purposeList != null && purposeList.Count > 0
    ? string.Join("", purposeList.Select((p, i) =>
        $"<div class='t-14 tab2'>{p.Detail}</div>"))
    : "")}  
<P class='t-12 tab2'><b>ข้อ 1 ขอบเขตความร่วมมือของ “สสว.”</b></P>
    <P class='t-12 tab3'>1.1 ตกลงร่วมดำเนินการโครงการโดยสนับสนุนงบประมาณ จำนวน {result.Contract_Value?.ToString("N2") ?? "0.00"} บาท  ( {strContract_Value} ) ซึ่งได้รวมภาษีมูลค่าเพิ่ม ตลอดจนค่าภาษีอากรอื่น ๆ แล้วให้กับ “{result.OrgName ?? ""}” และการใช้จ่ายเงินให้เป็นไปตามแผน การจ่ายเงินตามเอกสารแนบท้ายบันทึกข้อตกลงฉบับนี้</P>
    <P class='t-12 tab3'>1.2 ประสานการดำเนินโครงการ เพื่อให้บรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์</P>
    <P class='t-12 tab3'>1.3 กำกับ ติดตามและประเมินผลการดำเนินงานของโครงการ</P>
    <P class='t-12 tab2'><b>ข้อ 2 ขอบเขตความร่วมมือของ “{result.OrgName ?? ""}”</b></P>
    <P class='t-12 tab3'>2.1 ตกลงที่จะร่วมดำเนินการโครงการตามวัตถุประสงค์ของการโครงการ 
และขอบเขต การดำเนินการตามรายละเอียดโครงการ แผนการดำเนินการ และแผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) ที่แนบท้ายบันทึกข้อตกลงฉบับนี้</P>
    <P class='t-12 tab3'>2.2 ต้องดำเนินโครงการ ปฏิบัติตามแผนการดำเนินงาน แผนการใช้จ่ายเงิน (หรืออาจมีคู่มือ การดำเนินโครงการก็ได้) อย่างเคร่งครัดและให้แล้วเสร็จภายในระยะเวลาโครงการ</P>
    <P class='t-12 tab3'>2.3 ต้องประสานการดำเนินโครงการ เพื่อให้โครงการบรรลุวัตถุประสงค์ เป้าหมายผลผลิต และผลลัพธ์</P>
    <P class='t-12 tab3'>2.4 ต้องให้ความร่วมมือกับ สสว.ในการกำกับ ติดตามและประเมินผลการ ดำเนินงานของโครงการ</P>
    <P class='t-12 tab2'><b>ข้อ 3 อื่น ๆ</b></P>
    <P class='t-12 tab3'>3.1 หากฝ่ายใดฝ่ายหนึ่งประสงค์จะขอแก้ไข เปลี่ยนแปลง ขยายระยะเวลาของโครงการ จะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษร และต้องได้รับความยินยอมเป็นลาย ลักษณ์อักษรจากอีกฝ่ายหนึ่ง และต้องทำบันทึกข้อตกลงแก้ไข เปลี่ยนแปลง ขยายระยะเวลา เพื่อลงนาม ยินยอมทั้งสองฝ่าย</P>
   
<P class='t-12 tab3'>3.2 หากฝ่ายใดฝ่ายหนึ่งประสงค์จะขอบอกเลิกบันทึกข้อตกลงความร่วมมือก่อนครบ 
กำหนดระยะเวลาดำเนินโครงการจะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษรไม่น้อยกว่า 30 วัน และต้องได้รับความยินยอมเป็นลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และ “{result.OrgName ?? ""}” จะต้องคืนเงินในส่วน ที่ยังไม่ได้ใช้จ่ายหรือส่วนที่เหลือทั้งหมดพร้อมดอกผล (ถ้ามี) ให้แก่ สสว. ภายใน 15 วัน นับจากวันที่ได้รับ หนังสือของฝ่ายที่ยินยอมให้บอกเลิก</P>
 
<P class='t-12 tab3'>3.3 สสว. อาจบอกเลิกบันทึกข้อตกลงความร่วมมือได้ทันที หากตรวจสอบ หรือปรากฏ ข้อเท็จจริงว่า การใช้จ่ายเงินของ “{result.OrgName ?? ""}” ไม่เป็นไปตามวัตถุประสงค์ของโครงการ แผนการดำเนินงาน และแผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) ทั้งมีสิทธิเรียกเงินคงเหลือพร้อมดอกผล (ถ้ามี) คืนทั้งหมดได้ทันที</P>
    <P class='t-12 tab3'>3.4 ทรัพย์สินใด ๆ และ/หรือ สิทธิใด ๆ ที่ได้มาจากเงินสนับสนุนตามบันทึกข้อตกลงฉบับนี้ เมื่อสิ้นสุดโครงการให้ตกได้แก่ สสว. ทั้งสิ้น เว้นแต่ สสว. จะกำหนดให้เป็นอย่างอื่น</P>
    <P class='t-12 tab3'>3.5 “{result.OrgName ?? "-"}” ต้องไม่ดำเนินการในลักษณะการจ้างเหมา กับหน่วยงาน องค์กร หรือบุคคลอื่น ๆ ยกเว้นกรณีการจัดหา จัดจ้าง เป็นกิจกรรมหรือเป็นเรื่อง ๆ</P>
    <P class='t-12 tab3'>3.6 ในกรณีที่การดำเนินการตามบันทึกข้อตกลงฉบับนี้ เกี่ยวข้องกับข้อมูลส่วนบุคคล และ การคุ้มครองทรัพย์สินทางปัญญา “{result.OrgName ?? "-"}” จะต้องปฏิบัติตามกฎหมายว่าด้วยการคุ้มครองข้อมูล ส่วนบุคคลและการคุ้มครองทรัพย์สินทางปัญญาอย่างเคร่งครัด และหากเกิดความเสียหายหรือมีการฟ้อง ร้องใดๆ “{result.OrgName ?? "-"}” จะต้องเป็นผู้รับผิดชอบต่อการละเมิดบทบัญญัติแห่งกฎหมายดังกล่าว แต่เพียงฝ่ายเดียวโดยสิ้นเชิง</P>
    <P class='t-12 tab3'>บันทึกข้อตกลงนี้ทำขึ้นเป็นบันทึกข้อตกลงอิเล็กทรอนิกส์ คู่ตกลงได้อ่าน เข้าใจเงื่อนไข และยอมรับเงื่อนไข และได้ยืนยันว่าเป็นผู้มีอำนาจลงนามในบันทึกข้อตกลงจึงได้ลงลายมืออิเล็กทรอนิกส์พร้อมทั้งประทับตรา (ถ้ามี) ในบันทึกข้อตกลงไว้ และต่างฝ่ายต่างยึดถือไว้ฝ่ายละหนึ่งฉบับในระบบของตน</P>


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
    #endregion
}
