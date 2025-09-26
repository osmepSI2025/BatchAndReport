using BatchAndReport.DAO;
using BatchAndReport.Models;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

public class WordEContract_DataSecretService
{
    private readonly WordServiceSetting _w;
    E_ContractReportDAO _eContractReportDAO;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    public WordEContract_DataSecretService(WordServiceSetting ws
         , E_ContractReportDAO eContractReportDAO
          , IConverter pdfConverter
        )
    {
        _w = ws;
        _eContractReportDAO = eContractReportDAO;
         _pdfConverter = pdfConverter;
    }
    #region  4.1.1.2.7.สัญญาการรักษาข้อมูลที่เป็นความลับ
  

    public async Task<string> OnGetWordContact_DataSecretService_ToPDF(string id,string typeContact) 
    {
        var result = await _eContractReportDAO.GetNDAAsync(id);
        var conPurpose = await _eContractReportDAO.GetNDA_RequestPurposeAsync(id);
        
        var purposeHtml = "";
        if (conPurpose != null && conPurpose.Count > 0)
        {
            purposeHtml += "<p class='tab2 t-12'><b>ข้อ ๑ วัตถุประสงค์</b></p>";
            purposeHtml += "<p class='tab2 t-12'>โดยที่ผู้ให้ข้อมูลเป็นเจ้าของข้อมูล ผู้รับข้อมูลมีความต้องการที่จะใช้ข้อมูลของผู้ให้ข้อมูลเพื่อที่จะดำเนินการตามวัตถุประสงค์ ดังนี้ </p>";
            foreach (var purpose in conPurpose)
            {
                purposeHtml += $"<p class='tab2 t-12'>{System.Net.WebUtility.HtmlEncode(CommonDAO.ConvertStringArabicToThaiNumerals(purpose.Detail))}</p>";
            }
        }
        
        var conConfidentialHtml = "";
        var conConfidentialType = await _eContractReportDAO.GetNDA_ConfidentialTypeAsync(id);
        if (conConfidentialType != null && conConfidentialType.Count > 0)
        {
            foreach (var confidential in conConfidentialType)
            {
                conConfidentialHtml += $"<div class='tab2 t-12'>{System.Net.WebUtility.HtmlEncode(CommonDAO.ConvertStringArabicToThaiNumerals(confidential.Detail))}</div>";
            }
        }
        if (result == null)
        {
            throw new Exception("NDA data not found.");
        }

        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabun.ttf");
        string fontBase64 = "";
        if (File.Exists(fontPath))
        {
            var bytes = File.ReadAllBytes(fontPath);
            fontBase64 = Convert.ToBase64String(bytes);
        }
        // Read CSS file content
        var cssPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "css", "contract.css");
        string contractCss = "";
        if (File.Exists(cssPath))
        {
            contractCss = File.ReadAllText(cssPath, Encoding.UTF8);
        }
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
             <img src='data:image/jpeg;base64,{contractlogoBase64}'  height='80' />
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
        string strAttorneyLetterDate = CommonDAO.ToThaiDateStringCovert(result.Grant_Date ?? DateTime.Now);
        string strAttorneyLetterDate_CP = CommonDAO.ToThaiDateStringCovert(result.CP_S_AttorneyLetterDate ?? DateTime.Now);
        string strAttorneyOsmep = "";
        var HtmlAttorneyOsmep = new StringBuilder();
        if (result.AttorneyFlag == true)
        {
            strAttorneyOsmep = "ผู้มีอำนาจกระทำการแทนปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ เลขคำสั่งสำนักงานที่ " + result.AttorneyLetterNumber + " ฉบับลง" + strAttorneyLetterDate + "";

        }
        else
        {
            strAttorneyOsmep = "";
        }
        string strAttorney = "";
        var HtmlAttorney = new StringBuilder();
        if (result.CP_S_AttorneyFlag == true)
        {
            strAttorney = "ผู้มีอำนาจ กระทำการแทน ปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ ฉบับลง" + strAttorneyLetterDate_CP + "";

        }
        else
        {
            strAttorney = "";
        }
        #endregion

        var strDateTH = CommonDAO.ToThaiDateStringCovert(result.Sign_Date ?? DateTime.Now);


        #region signlist 

        var signlist = await _eContractReportDAO.GetSignNameAsync(id, typeContact);
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


        var html = $@"<html>
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
</head><body>
    <!-- Top Row: Logo left, Contract code box right -->
<table style='width:100%; border-collapse:collapse; margin-top:40px;'>
    <tr>
        <!-- Left: SME logo -->
        <td style='width:60%; text-align:left; vertical-align:top;'>
        <div style='display:inline-block;  padding:20px; font-size:32pt;'>
             <img src='data:image/jpeg;base64,{logoBase64}'  height='80' />
            </div>
        </td>
        <!-- Right: Contract code box (replace with your actual contract code if needed) -->
       <td style='width:40%; text-align:center; vertical-align:top;'>
            {contractLogoHtml}
        </td>
    </tr>
</table>
</br>
    <!-- Titles -->
    <div class='text-center t-14'><B>สัญญาการรักษาข้อมูลที่เป็นความลับ</B></div>
    <div class='text-center t-14'><B>(Non-disclosure Agreement : NDA)</B></div>
    <div class='text-center t-14'><B>ระหว่าง</B></div>
    <div class='text-center t-14'><B>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B></div>
    <div class='text-center t-14'><B>กับ {result.Contract_Party_Name}</B></div>
    <div class='text-center  t-14'>---------------------------------------------</div>
</br>
    <!-- Main contract body -->
   <p class='tab2 t-12'>
        สัญญาการรักษาข้อมูลที่เป็นความลับ (“สัญญา”) ฉบับนี้จัดขึ้น เมื่อ {strDateTH} ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.) ระหว่าง
    </p>
   <p class='tab2 t-12'>
        สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B>  โดย {CommonDAO.ConvertStringArabicToThaiNumerals(result.OSMEP_NAME)} ตำแหน่ง {CommonDAO.ConvertStringArabicToThaiNumerals(result.OSMEP_POSITION)} {CommonDAO.ConvertStringArabicToThaiNumerals(
strAttorneyOsmep)} สำนักงานตั้งอยู่เลขที่ ๑๒๐ หมู่ ๓ ศูนย์ราชการเฉลิมพระเกียรติ ๘๐ พรรษา ๕ ธันวาคม ๒๕๕๐ (อาคารซี) ชั้น ๒, ๑๐, ๑๑ ถนนแจ้งวัฒนะ แขวงทุ่งสองห้อง เขตหลักสี่ กรุงเทพ ๑๐๒๑๐ ซึ่งต่อไปในสัญญานี้จะเรียกว่า “ผู้เปิดเผยข้อมูล”
        </p>
</br> 
<p class='tab2 t-12'>กับ {CommonDAO.ConvertStringArabicToThaiNumerals(result.Contract_Party_Name)} โดย {CommonDAO.ConvertStringArabicToThaiNumerals(result.CP_S_NAME)} ตำแหน่ง {CommonDAO.ConvertStringArabicToThaiNumerals(result.CP_S_POSITION)} {CommonDAO.ConvertStringArabicToThaiNumerals(strAttorney)} 
     สำนักงานตั้งอยู่เลขที่ {CommonDAO.ConvertStringArabicToThaiNumerals(result.OfficeLoc)} ซึ่งต่อไปในสัญญานี้จะเรียกว่า “ผู้รับข้อมูล”
    </p>
   <p class='tab2 t-12'>คู่สัญญาได้ตกลงทำสัญญากันมีข้อความดังต่อไปนี้</p>

    <!-- NDA Clauses -->
    <!-- Purpose -->
   {purposeHtml}

   <p class='tab2 t-12'>โดยผู้รับข้อมูลประสงค์ให้ผู้ให้ข้อมูลเปิดเผยข้อมูลแก่ผู้รับข้อมูลอย่างเป็นความลับทั้งก่อน
หรือหลังจากวันที่สัญญาฉบับนี้มีผลใช้บังคับดังที่ระบุไว้ข้างต้น โดยผู้ให้ข้อมูลมีความจำเป็นต้องเปิดเผยข้อมูล
ที่เป็นความลับของผู้ให้ข้อมูล เพื่อผู้รับข้อมูลจะได้นำข้อมูลดังกล่าวไปประกอบการจัดทำ {CommonDAO.ConvertStringArabicToThaiNumerals(
result.Contract_Party_Name)} ร่วมกัน โดยผู้ให้ข้อมูลประสงค์
ให้ผู้รับข้อมูลเก็บรักษาความลับไว้ภายใต้สัญญานี้
</p>

    <!-- Confidential Types -->
   <p class='tab2 t-12'><b>ข้อ ๒ ข้อมูลที่เป็นความลับ</b></p>
   <p class='tab2 t-12'><b>“ข้อมูลที่เป็นความลับ”</b> หมายความว่า บรรดาข้อความเอกสารข้อมูลตลอดจนรายละเอียดทั้ง
ปวงที่เป็นของผู้ให้ข้อมูล รวมถึงที่อยู่ในความครอบครองหรือควบคุมดูแลของผู้ให้ข้อมูล และไม่เป็นที่รับรู้ของ
สาธารณชนโดยทั่วไปไม่ว่าจะในรูปแบบที่จับต้องได้หรือไม่ก็ตาม หรือสื่อแบบใดไม่ว่าจะถูกดัดแปลงแก้ไขโดย
ผู้รับข้อมูลหรือไม่ และไม่ว่าจะเปิดเผยเมื่อใดและอย่างไร ให้ถือว่าเป็นความลับโดยข้อมูลที่เป็นความลับอาจ
หมายความรวมถึง ข้อมูลเชิงกลยุทธ์ของผู้ให้ข้อมูล แผนธุรกิจ ข้อมูลทางการเงิน ข้อมูลลูกจ้าง ข้อมูลผู้
ประกอบการ และข้อมูลส่วนบุคคลที่ผู้ให้ข้อมูลได้เก็บ รวบรวม ใช้ ข้อมูลที่เป็นความลับที่ผู้ให้ข้อมูล หรือในนามของผู้ให้ข้อมูลที่เปิดเผยแก่ผู้รับข้อมูล ซึ่งหมายความรวมถึงข้อมูลที่ผู้ให้ข้อมูลให้แก่ผู้รับข้อมูล ดังนี้
    </p>
   <p class='tab2 t-12'>(ระบุประเภทของข้อมูลที่เป็นความลับที่นำส่งให้แก่กัน)</p>

    {conConfidentialHtml}

    <!-- Clause 3 -->
   <p class='tab2 t-12'><b>ข้อ ๓ การรักษาข้อมูลที่เป็นความลับ</b></p>
   <p class='tab2 t-12'>๓.๑ ผู้รับข้อมูลต้องรับผิดชอบรักษาข้อมูลที่เป็นความลับและเก็บข้อมูลความลับไว้โดยครบ
ถ้วนและอย่างเคร่งครัด ผู้รับข้อมูลจะต้องไม่เปิดเผยทำสำเนาหรือทำการอื่นใดทำนองเดียวกันแก่บุคคลอื่น
ไม่ว่าทั้งหมดหรือบางส่วน เว้นแต่ได้รับอนุญาตเป็นหนังสือจากผู้ให้ข้อมูล</p>

   <p class='tab2 t-12'>๓.๒ ผู้รับข้อมูลต้องใช้ข้อมูลที่เป็นความลับเพื่อการอันเกี่ยวกับหรือสัมพันธ์ 
กับการดำเนินงานที่มีอยู่ระหว่างผู้ให้ข้อมูลกับผู้รับข้อมูล โดยผู้รับข้อมูลต้องแจ้งให้ผู้ให้ข้อมูลทราบโดยทันที
ที่พบการใช้หรือการเปิดเผยข้อมูลที่เป็นความลับโดยไม่ได้รับอนุญาตหรือการละเมิดหรือฝ่าฝืนข้อกำหนด
ตามสัญญานี้ อีกทั้ง ผู้รับข้อมูลจะต้องให้ความร่วมมือกับผู้ให้ข้อมูลอย่างเต็มที่ในการเรียกคืนซึ่งการ 
ครอบครองข้อมูลที่เป็นความลับการป้องกันการใช้ข้อมูลที่เป็นความลับโดยไม่ได้รับอนุญาตและการระงับ 
ยับยั้งการเผยแพร่ข้อมูลที่เป็นความลับออกสู่สาธารณะ</p>
   <p class='tab2 t-12'>๓.๓ ผู้รับข้อมูลต้องจัดให้มีและคงไว้ซึ่งมาตรการรักษาความปลอดภัยสำหรับ 
การจัดเก็บและประมวลผลข้อมูลที่มีความเหมาะสมในมาตรการเชิงองค์กร มาตรการเชิงเทคนิค และ 
มาตรการเชิงกายภาพโดยคำนึงถึงลักษณะ ขอบเขต และวัตถุประสงค์ของการดำเนินการตามวัตถุประสงค์ 
ที่ของสัญญาฉบับนี้เป็นสำคัญ เพื่อป้องกันมิให้ข้อมูลที่เป็นความลับถูกนำไปใช้โดยมิได้รับอนุญาตหรือ 
ถูกเปิดเผยแก่บุคคลอื่น โดยผู้รับข้อมูลต้องใช้มาตรการการเก็บรักษาข้อมูลที่เป็นความลับในระดับเดียวกัน 
กับที่ผู้รับข้อมูลใช้กับข้อมูลที่เป็นความลับของตนเองซึ่งต้องไม่น้อยกว่าการดูแลที่สมควร</p>
   <p class='tab2 t-12'>๓.๔ ผู้รับข้อมูลต้องแจ้งให้บุคลากร พนักงาน ลูกจ้าง ที่ปรึกษาของผู้รับข้อมูล 
และ/หรือบุคคลภายนอกที่ต้องเกี่ยวข้องกับข้อมูลที่เป็นความลับนั้น ทราบถึงความเป็นความลับและ 
ข้อจำกัดสิทธิในการใช้และการเปิดเผยข้อมูลที่เป็นความลับ และผู้รับข้อมูลต้องดำเนินการให้บุคคลดังกล่าว 
ต้องผูกพันด้วยสัญญาหรือข้อตกลงเป็นหนังสือในการรักษาข้อมูลที่เป็นความลับ โดยมีข้อกำหนดเช่น 
เดียวกับหรือไม่น้อยกว่าข้อกำหนดและเงื่อนไขในสัญญาฉบับนี้ด้วย</p>
   <p class='tab2 t-12'>๓.๕ ข้อมูลที่เป็นความลับตามสัญญาฉบับนี้ไม่รวมไปถึงข้อมูลดังต่อไปนี้</p>
   <p class='tab3 t-12'>(๑) ข้อมูลที่ผู้ให้ข้อมูลเปิดเผยแก่สาธารณะ</p>
   <p class='tab3 t-12'>(๒) ข้อมูลที่ผู้รับข้อมูลทราบอยู่ก่อนที่ผู้ให้ข้อมูลจะเปิดเผยข้อมูลนั้น</p>
   <p class='tab3 t-12'>(๓) ข้อมูลที่มาจากการพัฒนาโดยอิสระของผู้รับข้อมูลเอง</p>
   <p class='tab3 t-12'>(๔) ข้อมูลที่ต้องเปิดเผยโดยกฎหมายหรือตามคำสั่งศาล ทั้งนี้ ผู้รับข้อมูลต้องมีหนัง
สือแจ้งผู้ให้ข้อมูลได้รับทราบถึงข้อกำหนดหรือคำสั่งดังกล่าว โดยแสดงเอกสารข้อกำหนด หมายศาลและ/
หรือหมายค้นอย่างเป็นทางการต่อผู้ให้ข้อมูลก่อนที่จะดำเนินการเปิดเผยข้อมูลดังกล่าว และในการเปิดเผย
ข้อมูลดังกล่าวผู้รับข้อมูลจะต้องดำเนินการตามขั้นตอนทางกฎหมายเพื่อขอให้คุ้มครองข้อมูลดังกล่าวไม่ให้
ถูกเปิดเผยต่อสาธารณะด้วย</p>
   <p class='tab3 t-12'>(๕) ผู้รับข้อมูลได้รับความยินยอมเป็นลายลักษณ์อักษรให้เปิดเผยข้อมูลจากผู้ให้ข้อมูล
ก่อนที่ผู้รับข้อมูลจะเปิด เผยข้อมูลนั้น</p>
   <p class='tab3 t-12'>(๖) ผู้รับข้อมูลได้รับข้อมูลที่เป็นความลับจากบุคคลที่สามที่ไม่อยู่ภายใต้ข้อกำหนดใน
เรื่องการรักษาความลับ หรือข้อจำกัดในเรื่องสิทธิ</p>
   <p class='tab2 t-12'>๓.๖ ผู้รับข้อมูลต้องไม่ทำซ้ำข้อมูลที่เป็นความลับแม้เพียงส่วนหนึ่งส่วนใดหรือทั้งหมด
เว้นแต่การทำซ้ำเพื่อการใช้ข้อมูลที่เป็นความลับให้บรรลุผลตามวัตถุประสงค์ที่กำหนดไว้ในสัญญานี้ และ
ไม่ทำวิศวกรรมย้อนกลับ หรือถอดรหัสข้อมูลที่เป็นความลับ ต้นแบบ หรือสิ่งอื่นใดที่บรรจุข้อมูลที่เป็น
ความลับ รวมทั้งไม่เคลื่อนย้าย พิมพ์ทับ หรือทำให้เสียรูปซึ่งสัญลักษณ์ที่แสดงเครื่องหมายสิทธิบัตร
อนุสิทธิบัตร ลิขสิทธิ์ เครื่องหมายการค้า ตราสัญลักษณ์ และเครื่องหมายอื่นใดที่แสดงกรรมสิทธิ์ของ
ต้นแบบหรือสำเนาของข้อมูลที่เป็นความลับที่ได้รับมาจากผู้ให้ข้อมูล</p>

    <!-- Clause 4 -->
   <p class='tab2 t-12'><b>ข้อ ๔ ทรัพย์สินทางปัญญา</b></p>
   <p class='tab2 t-12'>สัญญาฉบับนี้ไม่มีผลบังคับใช้เป็นการโอนสิทธิหรือการอนุญาตให้ใช้สิทธิ (ไม่ว่าโดยตรง
 หรือโดยอ้อม) ให้แก่ผู้รับข้อมูลที่ได้รับความลับซึ่งสิทธิบัตร ลิขสิทธิ์ การออกแบบ เครื่องหมายการค้าตรา
สัญลักษณ์ รูปประดิษฐ์อื่นใดชื่อทางการค้า ความลับทางการค้า ไม่ว่าจดทะเบียนไว้ตามกฎหมายหรือไม่
ก็ตามหรือสิทธิอื่น ๆ ของผู้ให้ข้อมูล ซึ่งอาจปรากฏอยู่หรือนำมาทำซ้ำไว้ในเอกสารข้อมูลที่เป็นความลับทั้งนี้
ผู้รับข้อมูลหรือบุคคลอื่นใดที่เกี่ยวข้องกับผู้รับข้อมูลและเกี่ยวข้องกับข้อมูลที่เป็นความลับดังกล่าวจะไม่ยื่น
ขอรับสิทธิและ/หรือขอจดทะเบียนเกี่ยวกับทรัพย์สินทางปัญญาใด ๆ ตลอดจนไม่นำไปใช้โดยไม่ได้รับ
การอนุญาตเป็นลายลักษณ์อักษรจากผู้ให้ข้อมูลเกี่ยวกับรายละเอียดข้อมูลที่เป็นความลับหรือส่วนหนึ่งส่วนใด
ของรายละเอียดดังกล่าว</p>

    <!-- Clause 5 -->
   <p class='tab2 t-12'><b>ข้อ ๕ การส่งคืน ลบ หรือการทำลายข้อมูลที่เป็นความลับ</b></p>
   <p class='tab2 t-12'>เมื่อการดำเนินงานที่มีอยู่ระหว่างผู้ให้ข้อมูลกับผู้รับข้อมูลเสร็จสิ้นลงตามวัตถุประสงค์ผู้รับ
ข้อมูลจะต้องส่งมอบข้อมูลที่เป็นความลับและสำเนาของข้อมูลที่เป็นความลับที่ผู้รับข้อมูลได้รับไว้คืนให้แก่ผู้ให้
ข้อมูล เว้นแต่ผู้ให้ข้อมูลเห็นว่าไม่ต้องนำส่งคืนแต่ต้องเลิกใช้ข้อมูลที่เป็นความลับ และทำการลบหรือทำลาย
ข้อมูลที่เป็นความลับทั้งถูกจัดเก็บไว้ในคอมพิวเตอร์หรืออุปกรณ์อื่นใดที่ใช้จัดเก็บข้อมูล (ถ้ามี) หรือดำเนิน
การอื่นตามที่ได้รับการแจ้งเป็นลายลักษณ์อักษรจากผู้ให้ข้อมูล</p>

    <!-- Clause 6 -->
   <p class='tab2 t-12'><b>ข้อ ๖ การชดใช้ค่าเสียหาย</b></p>
   <p class='tab2 t-12'>๖.๑ กรณีที่ผู้รับข้อมูล พนักงาน ลูกจ้าง ที่ปรึกษาของผู้รับข้อมูล และ/หรือบุคคลภาย
นอกที่ได้รับข้อมูลที่เป็นความลับจากผู้รับข้อมูลฝ่าฝืนข้อกำหนดตามสัญญานี้และก่อให้เกิดความเสียหายแก่ผู้
ให้ข้อมูล และ/หรือบุคคลอื่นผู้รับข้อมูลจะต้องชดใช้ค่าเสียหายให้แก่ผู้ให้ข้อมูล และ/หรือบุคคลที่ได้รับความ
เสียหายสำหรับความเสียหายเช่นว่านั้น ทั้งนี้ ผู้รับข้อมูลจะต้องแจ้งแก่ผู้ให้ข้อมูลทราบเป็นลายลักษณ์อักษร
ภายใน ๗ วันนับตั้งแต่มีการละเมิดข้อมูลที่เป็นความลับเกิดขึ้น</p>

   <p class='tab2 t-12'>๖.๒ ผู้รับข้อมูลรับทราบว่าการเปิดเผยหรือการใช้ข้อมูลที่เป็นความลับโดยฝ่าฝืนข้อ
กำหนดตามสัญญานี้จะก่อให้เกิดความเสียหายแก่ผู้ให้ข้อมูลในจำนวนที่ไม่สามารถประเมินได้ดังนั้นผู้รับ
ข้อมูลยินยอมให้ผู้ให้ข้อมูลใช้สิทธิที่จะร้องขอต่อศาลเพื่อให้มีคำสั่งให้ผู้รับข้อมูลหยุดการกระทำใดๆที่เป็น
การฝ่าฝืนข้อกำหนดตามสัญญานี้ และ/หรือใช้วิธีคุ้มครองชั่วคราวใดๆ ตามที่ผู้ให้ข้อมูลเห็นว่าเหมาะสม
ได้โดยผู้รับข้อมูลจะเป็นผู้รับผิดชอบค่าใช้จ่ายต่าง ๆ ที่เกิดขึ้นทั้งหมดจากการดำเนินการดังกล่าว</p>
   <p class='tab2 t-12'>๖.๓ กรณีที่ผู้ให้ข้อมูลสงสัยว่าผู้รับข้อมูลฝ่าฝืนข้อกำหนดตามสัญญานี้ ผู้รับข้อมูล 
จะต้องเป็นฝ่ายพิสูจน์ว่าผู้รับข้อมูลไม่ได้ฝ่าฝืนข้อกำหนดตามสัญญานี้</p>

    <!-- Clause 7 -->
   <p class='tab2 t-12'><b>ข้อ ๗ ระยะเวลาตามสัญญา</b></p>
   <p class='tab2 t-12'>สัญญานี้มีผลบังคับใช้นับตั้งแต่วันที่ทำสัญญานี้ โดยมีกำหนดระยะเวลาทั้งสิ้น {(result.EnforcePeriods.HasValue ? CommonDAO.ConvertStringArabicToThaiNumerals(result.EnforcePeriods.Value.ToString()) : "-")} ปี นับตั้งแต่

วันที่ทำสัญญาฉบับนี้</p>
   <p class='tab2 t-12'>เมื่อครบกำหนดระยะเวลาตามวรรคหนึ่ง หรือเมื่อมีการบอกเลิกสัญญา หรือผู้ให้ข้อมูล
ได้แจ้งให้ผู้รับข้อมูลดำเนินการทำลายข้อมูลดังกล่าว ผู้รับข้อมูลจะต้องดำเนินการทำลายข้อมูล ภายใน ๗
วันนับแต่ได้รับหนังสือร้องขอจากผู้ให้ข้อมูล ทั้งนี้ ผู้รับข้อมูลจะต้องไม่มีการสงวนไว้ซึ่งสำเนาใด ๆ</p>

    <!-- Clause 8 -->
   <p class='tab2 t-12'><b>ข้อ ๘ ข้อตกลงอื่น ๆ</b></p>
   <p class='tab2 t-12'>๘.๑ ในกรณีที่มีเหตุจำเป็นต้องมีการเปลี่ยนแปลงแก้ไขสัญญานี้ ให้ทำเป็นลายลักษณ์อักษร
และลงนามโดยคู่สัญญาหรือผู้มีอำนาจลงนามผูกพันนิติบุคคลและประทับตราสำคัญของนิติบุคคล (ถ้ามี) ของคู่สัญญา แล้วแต่กรณี</p>
   <p class='tab2 t-12'>๘.๒ กรณีที่ผู้รับข้อมูลได้โอนกิจการ รวมกิจการ หรือควบกิจการ หรือดำเนินการอื่น ๆ ใน
ลักษณะที่มีการเปลี่ยนแปลงของวัตถุประสงค์ในการดำเนินกิจการของผู้รับข้อมูลผู้รับข้อมูลจะต้องแจ้งให้ผู้ให้
ข้อมูลทราบภายใน ๕ วันทำการ นับแต่ได้เกิดเหตุดังกล่าวขึ้น</p>

    <!-- Clause 9 -->
   <p class='tab2 t-12'><b>ข้อ ๙ การบังคับใช้</b></p>
   <p class='tab2 t-12'>๙.๑ ในกรณีที่ปรากฏภายหลังว่าส่วนใดส่วนหนึ่งในสัญญาฉบับนี้เป็นโมฆะให้ถือว่า 
ข้อกำหนดส่วนที่เป็นโมฆะไม่มีผลบังคับในสัญญานี้ และข้อกำหนดที่เหลืออยู่ในสัญญาฉบับนี้ยังคงใช้บังคับ 
และมีผลอยู่อย่างสมบูรณ์</p>
   <p class='tab2 t-12'>๙.๒ สัญญาฉบับนี้อยู่ภายใต้การบังคับและตีความตามกฎหมายของประเทศไทย ให้ศาลของ
ประเทศไทยมีอำนาจในกรณีที่มีข้อพิพาทใด ๆ อันเกิดขึ้นจากสัญญาฉบับนี้</p>
   <p class='tab2 t-12'>สัญญานี้ทำขึ้นเป็นสัญญาอิเล็กทรอนิกส์ คู่สัญญาได้อ่าน เข้าใจเงื่อนไข และยอมรับเงื่อนไข และได้ยืนยันว่าเป็นผู้มีอำนาจลงนามในสัญญาจึงได้ลงลายมืออิเล็กทรอนิกส์พร้อมทั้งประทับตรา (ถ้ามี) ในสัญญาไว้ และต่างฝ่ายต่างยึดถือไว้ฝ่ายละหนึ่งฉบับในระบบของตน</p>

    <!-- Signature Table -->
   </br>
</br>
{signatoryTableHtml}
    <P class='t-12 tab2'>ข้าพเจ้าขอรับรองว่า ทั้งสองฝ่ายได้ลงนามในสัญญาโดยวิธีการอิเล็กทรอนิกส์ เพื่อแสดงเจตนาของคู่สัญญาแล้ว ข้าพเจ้าจึงได้ลงลายมือชื่ออิเล็กทรอนิกส์รับรองเป็นพยานในสัญญาพร้อมนี้</P>

{signatoryTableHtmlWitnesses}
</body>
</html>
";


        return html;
    }
    #endregion  4.1.1.2.7.สัญญาการรักษาข้อมูลที่เป็นความลับ
}
