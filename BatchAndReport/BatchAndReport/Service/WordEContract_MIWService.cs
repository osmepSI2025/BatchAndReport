using BatchAndReport.DAO;
using BatchAndReport.Models;
using BatchAndReport.Services;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Commons.Bouncycastle.Crypto;
using iText.Signatures;
using SkiaSharp;
using Spire.Doc.Documents;
using System.Text;
using System.Threading.Tasks;
using static SkiaSharp.HarfBuzz.SKShaper;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
public class WordEContract_MIWService
{
    private readonly WordServiceSetting _w;
    private readonly E_ContractReportDAO _eContractReportDAO;
    private readonly Econtract_Report_MIWDAO _MIWDAO;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    private readonly EContractDAO _eContractDAO;

    public WordEContract_MIWService(
        WordServiceSetting ws,
        E_ContractReportDAO eContractReportDAO,
         Econtract_Report_MIWDAO  _Report_MIWDAO
      , IConverter pdfConverter
        ,
EContractDAO eContractDAO   
    )
    {
        _w = ws;
        _eContractReportDAO = eContractReportDAO;
        _pdfConverter = pdfConverter;
        _MIWDAO = _Report_MIWDAO;
        _eContractDAO = eContractDAO;
    }

    public async Task<string> OnGetWordContact_MIWServiceHtmlToPDF(string conId)
    {
        try {
            var dataResult = await _MIWDAO.GetMIWAsync(conId);
            if (dataResult == null)
                throw new Exception("MIW data not found.");
            var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabun.ttf");
            string fontBase64 = "";
            if (File.Exists(fontPath))
            {
                var bytes = File.ReadAllBytes(fontPath);
                fontBase64 = Convert.ToBase64String(bytes);
            }
            string contractCss = "";
            var cssPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "css", "contract.css");

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
            if (!string.IsNullOrEmpty(dataResult.Organization_Logo) && dataResult.Organization_Logo.Contains("<content>"))
            {
                try
                {
                    // ตัดเอาเฉพาะ Base64 ในแท็ก <content>...</content>
                    var contentStart = dataResult.Organization_Logo.IndexOf("<content>") + "<content>".Length;
                    var contentEnd = dataResult.Organization_Logo.IndexOf("</content>");
                    var contractlogoBase64 = dataResult.Organization_Logo.Substring(contentStart, contentEnd - contentStart);

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
            string strAttorneyLetterDate = CommonDAO.ToArabicDateStringCovert(dataResult.Grant_Date ?? DateTime.Now);
            string strAttorneyLetterDate_CP = CommonDAO.ToArabicDateStringCovert(dataResult.CP_S_AttorneyLetterDate ?? DateTime.Now);
            string strAttorneyOsmep = "";
            var HtmlAttorneyOsmep = new StringBuilder();
            if (dataResult.AttorneyFlag == true)
            {
                strAttorneyOsmep = "ผู้มีอำนาจกระทำการแทนปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ เลขคำสั่งสำนักงานที่ " + dataResult.AttorneyLetterNumber + " ฉบับลง" + strAttorneyLetterDate + "";

            }
            else
            {
                strAttorneyOsmep = "";
            }
            string strAttorney = "";
            var HtmlAttorney = new StringBuilder();
            if (dataResult.CP_S_AttorneyFlag == true)
            {
                strAttorney = "ผู้มีอำนาจ กระทำการแทน ปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ ฉบับลง" + strAttorneyLetterDate_CP + "";

            }
            else
            {
                strAttorney = "";
            }
            #endregion

            #region signlist 

            var signlist = await _eContractReportDAO.GetSignNameAsync(conId, "MIW");
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

            #region Document Attach
            var listDocAtt = await _eContractDAO.GetRelatedDocumentsAsync(conId, "MIW");
            var htmlDocAtt = listDocAtt != null
                ? string.Join("", listDocAtt.Select(docItem =>
                    $"<p class='tab3 t-14'>{docItem.DocumentTitle} จำนวน {docItem.PageAmount} หน้า</div>"))
                : "";
            #endregion

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



</br>
    <div class='t-12 text-center'><b>บันทึกข้อตกลง</b></div>
    <div class='t-12 text-center'><b>จ้างเหมาบริการ {dataResult.ServiceName} </b></div>
</br>
    <div class='t-12 text-right'><b>บันทึกข้อตกลงเลขที่ {dataResult.Contract_Number}</b></div>
 
</br>
    <P class='t-12 tab2'>
        บันทึกข้อตกลงฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เลขที่ ๑๒๐ หมู่ ๓ ศูนย์ราชการเฉลิมพระเกียรติ ๘๐ พรรษา ๕ ธันวาคม ๒๕๕๐ (อาคารซี) ชั้น ๒, ๑๐, ๑๑ ถนนแจ้งวัฒนะ แขวงทุ่งสองห้อง เขตหลักสี่ กรุงเทพมหานคร ๑๐๒๑๐ 
 เมื่อวันที่ {dataResult.ContractSignDate.Value.ToString("dd/MM/yyyy")} ระหว่าง สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม 
โดย {dataResult.OSMEP_NAME} ตำแหน่ง {dataResult.OSMEP_POSITION} {strAttorneyOsmep}  ซึ่งต่อไปในบันทึกข้อตกลงนี้เรียกว่า “ผู้ว่าจ้าง” ฝ่ายหนึ่ง 
กับ {dataResult.ContractPartyName} ผู้ถือบัตรประจำตัวประชาชนเลขที่ {dataResult.IdenID}
วันออกบัตร {dataResult.IdenIssue_Date.Value.ToString("dd/MM/yyyy")} บัตรหมดอายุ {dataResult.IdenExpiry_Date.Value.ToString("dd/MM/yyyy")} 
อยู่บ้านเลขที่ {dataResult.AddressNo} ถนน {dataResult.AddressStreet} 
ตำบล {dataResult.AddressSubDistrict} อำเภอ {dataResult.AddressDistrict} จังหวัด {dataResult.AddressProvince + " " + dataResult.AddressZipCode} 
ปรากฏตามเอกสารแนบท้ายบันทึกข้อตกลงนี้ ซึ่งต่อไปในบันทึกข้อตกลงนี้เรียกว่า “ผู้รับจ้าง” อีกฝ่ายหนึ่ง

    </P>

<p class='t-12 tab2'>ทั้งสองฝ่ายได้ตกลงกันมีข้อความดังต่อไปนี้</p>
<p class='t-12 tab2'><b>ข้อ ๑.ข้อตกลงว่าจ้าง</b></p>
  <P class='t-12 tab2'>จ้างและผู้รับจ้างตกลงรับจ้างเหมาบริการ {dataResult.HiringAgreement}
ตามข้อกำหนดและเงื่อนไขของบันทึกข้อตกลงนี้รวมทั้งเอกสารแนบท้ายบันทึกข้อตกลง
    </P>
    <P class='t-12 tab2'><b>ข้อ ๒.เอกสารแนบท้ายบันทึกข้อตกลง</b>
</P>
    <P class='t-12 tab3'>เอกสารแนบท้ายบันทึกข้อตกลงดังต่อไปนี้ให้ถือเป็นส่วนหนึ่งของบันทึกข้อตกลงนี้
</P>
{htmlDocAtt}
 <P class='t-12 tab3'>ความใดในเอกสารแนบท้ายบันทึกข้อตกลงที่ขัดแย้งกับข้อความในบันทึกข้อตกลงนี้ให้
ใช้ข้อความในบันทึกข้อตกลงนี้บังคับในกรณีที่เอกสารแนบท้ายบันทึกข้อตกลงขัดแย้งกันเอง
ผู้รับจ้างจะต้องปฏิบัติตามคำวินิจฉัยของผู้ว่าจ้าง
ทั้งนี้ ผู้รับจ้างไม่มีสิทธิเรียกร้องค่าเสียหายหรือค่าใช้จ่ายใดๆ ทั้งสิ้น
</P>
    <P class='t-12 tab2'><B>ข้อ ๓. ค่าจ้างและการจ่ายเงิน 
</B></P>
    <p class='t-12 tab3'>ผู้ว่าจ้างตกลงจ่ายและผู้รับจ้างตกลงรับเงินค่าจ้างจำนวนเงิน  {dataResult.HiringAmount??0} บาท( {CommonDAO.NumberToThaiText(dataResult.HiringAmount ?? 0)} )
ซึ่งได้รวมภาษีอากรอื่นๆ และค่าใช้จ่ายทั้งปวงด้วยแล้ว โดยกำหนดการจ่ายเงินเป็นรายเดือน จำนวนเงินเดือนละ {dataResult.Salary??0} บาท ( {CommonDAO.NumberToThaiText(dataResult.Salary ?? 0)} )
เมื่อผู้รับจ้างได้ปฏิบัติงานและนำส่งรายงานผลการปฏิบัติงาน และใบบันทึกลงเวลาการปฏิบัติงานในแต่ละเดือนให้แก่ผู้ว่าจ้าง 
ภายในวันที่ ๕ ของเดือนถัดไป นับจากวันสิ้นสุดของงานในแต่ละงวด ยกเว้นงวดสุดท้ายให้ส่งมอบภายในวันที่ {(dataResult.Delivery_Due_Date.HasValue ? dataResult.Delivery_Due_Date.Value.ToString("dd/MM/yyyy") : DateTime.Now.ToString("dd/MM/yyyy"))}
ซึ่งมีรายละเอียดของงานปรากฏตามเอกสารแนบท้ายบันทึกข้อตกลง และผู้ว่าจ้างได้ตรวจรับงานจ้างไว้โดยครบถ้วนแล้ว</p>
   <P class='t-12 tab3'>ทั้งนี้ หากเดือนใดมีการปฏิบัติงานไม่เต็มเดือนปฏิทิน ให้คิดค่าจ้างเหมาเป็นรายวัน ในอัตราวันละ {dataResult.DailyRate} บาท  ( {CommonDAO.NumberToThaiText(dataResult.DailyRate ?? 0)} ) </P>
    <P class='t-12 tab3'>การจ่ายเงินตามเงื่อนไขแห่งสัญญานี้ ผู้ว่าจ้างจะโอนเงินเข้าบัญชีเงินฝากธนาคาร ของผู้รับจ้าง 
ชื่อธนาคาร {dataResult.ContractBankName} สาขา {dataResult.ContractBankBranch} ชื่อบัญชี {dataResult.ContractBankAccountName}
เลขที่บัญชี {dataResult.ContractBankAccountNumber} ทั้งนี้ ผู้รับจ้างตกลงเป็นผู้รับภาระเงินค่าธรรมเนียมหรือค่าบริการอื่นใดเกี่ยวกับการโอน
รวมทั้งค่าใช้จ่ายอื่นใด (ถ้ามี) ที่ธนาคารเรียกเก็บ และยินยอมให้มีการหักเงินดังกล่าวจากจำนวนเงินโอนในงวด</P>

  <P class='t-12 tab2'><b>ข้อ ๔.กำหนดเวลาแล้วเสร็จและสิทธิของผู้ว่าจ้างในการบอกเลิกบันทึกข้อตกลง</b></P>
    <p class='t-12 tab3'>ผู้รับจ้างต้องเริ่มทำงานที่รับจ้างภายในวันที่  {dataResult.WorkStartDate.Value.ToString("dd/MM/yyyy")} และจะต้องทำงานให้แล้วเสร็จบริบูรณ์ภายในวันที่ {dataResult.WorkEndDate.Value.ToString("dd/MM/yyyy")}
ตามรายละเอียดเอกสารแนบท้ายบันทึกข้อตกลงนี้ และต้องผ่านการตรวจรับผลการปฏิบัติงานจากผู้ว่าจ้างในแต่ละเดือน
ถ้าผู้รับจ้างมิได้ลงมือทำงานภายในกำหนดเวลา หรือไม่สามารถทำงานให้ครบถ้วนตามเงื่อนไขของบันทึกข้อตกลงนี้
หรือมีเหตุให้เชื่อได้ว่า ผู้รับจ้างไม่สามารถทำงานให้แล้วเสร็จภายในกำหนดเวลา
หรือจะแล้วเสร็จล่าช้าเกินกว่ากำหนดเวลา หรือตกเป็นผู้ถูกพิทักษ์ทรัพย์เด็ดขาดหรือตกเป็นบุคคลล้มละลาย
หรือเพิกเฉยไม่ปฏิบัติตามคำสั่งของคณะกรรมการตรวจรับพัสดุ ผู้ว่าจ้างมีสิทธิที่จะบอกเลิกบันทึกข้อตกลงนี้ได้
และมีสิทธิจ้างผู้รับจ้างรายใหม่เข้าทำงานของผู้รับจ้างให้ลุล่วงไปได้ด้วย
การใช้สิทธิบอกเลิกบันทึกข้อตกลงนั้นไม่กระทบสิทธิของ
ผู้ว่าจ้างที่จะเรียกร้องค่าเสียหายจากผู้รับจ้าง</p>
    <p class='t-12 tab3'>การที่ผู้ว่าจ้างไม่ใช้สิทธิบอกเลิกบันทึกข้อตกลงดังกล่าวข้างต้นนั้น
ไม่เป็นเหตุให้ผู้รับจ้างพ้นจากความรับผิดตามบันทึกข้อตกลง</p>

<P class='t-12 tab2'><b>ข้อ ๕.การจ้างช่วง</b></P>
    <P class='t-12 tab3'>ผู้รับจ้างจะต้องไม่เอางานทั้งหมดหรือแต่บางส่วนของบันทึกข้อตกลงนี้ไปให้ผู้อื่นรับจ้างช่วงอีกทอดหนึ่ง
เว้นแต่การจ้างช่วงงานแต่บางส่วนที่ได้รับอนุญาตเป็นหนังสือจากผู้ว่าจ้างแล้ว การที่ผู้ว่าจ้างได้อนุญาตให้จ้างช่วงงานแต่บางส่วน 
ดังกล่าวนั้น ไม่เป็นเหตุให้ผู้รับจ้างหลุดพ้นจากความรับผิดหรือพันธะหน้าที่ตามบันทึกข้อตกลงนี้และผู้รับจ้างจะยังคงต้อง 
รับผิดในความผิดและความประมาทเลินเล่อของผู้รับจ้างช่วงหรือของตัวแทนหรือลูกจ้างของผู้รับจ้างช่วงนั้นทุกประการ
</P>
<P class='t-12 tab3'>กรณีผู้รับจ้างไปจ้างช่วงงานแต่บางส่วนโดยฝ่าฝืนความในวรรคหนึ่งผู้รับจ้างต้องชำระค่าปรับให้แก่ผู้ว่าจ้างเป็น 
จำนวนเงินในอัตราร้อยละ ๑๐.๐๐ (สิบ) ของวงเงินของงานที่จ้างช่วงตามบันทึกข้อตกลง ทั้งนี้ ไม่ตัดสิทธิผู้ว่าจ้างในการบอกเลิก บันทึกข้อตกลง</P>

<P class='t-12 tab2'><b>ข้อ ๖.ความรับผิดของผู้รับจ้าง</b></P>
<P class='t-12 tab3'>ผู้รับจ้างจะต้องรับผิดต่ออุบัติเหตุ ความเสียหาย หรือภยันตรายใดๆ อันเกิดจากการปฏิบัติงานของผู้รับจ้าง
และจะต้องรับผิดต่อความเสียหายจากการกระทำของลูกจ้างหรือตัวแทนของผู้รับจ้าง และจากการปฏิบัติงานของผู้รับจ้างช่วงด้วย (ถ้ามี)</P>
 <P class='t-12 tab3'>ความเสียหายใดๆ อันเกิดแก่งานที่ผู้รับจ้างได้ทำขึ้น แม้จะเกิดขึ้นเพราะเหตุสุดวิสัยก็ตาม 
ผู้รับจ้างจะต้องรับผิดชอบโดยซ่อมแซมให้คืนดีหรือเปลี่ยนให้ใหม่โดยค่าใช้จ่ายของผู้รับจ้างเอง เว้นแต่ความเสียหายนั้น เกิดจากความผิดของผู้ว่าจ้าง
ทั้งนี้ ความรับผิดของผู้รับจ้างดังกล่าวในข้อนี้จะสิ้นสุดลงเมื่อผู้ว่าจ้างได้รับมอบงานครั้งสุดท้าย</P>
<P class='t-12 tab3'>ผู้รับจ้างจะต้องรับผิดต่อบุคคลภายนอกในความเสียหายใดๆ อันเกิดจากการปฏิบัติงานของผู้รับจ้าง หรือลูกจ้าง 
หรือตัวแทนของผู้รับจ้าง รวมถึงผู้รับจ้างช่วง (ถ้ามี) ตามบันทึกข้อตกลงนี้ หากผู้ว่าจ้างถูกเรียกร้องหรือฟ้องร้อง 
หรือต้องชดใช้ค่าเสียหายให้แก่บุคคลภายนอกไปแล้ว ผู้รับจ้างจะต้องดำเนินการใดๆ เพื่อให้มีการว่าต่างแก้ต่างให้แก่ผู้ว่าจ้าง 
โดยค่าใช้จ่ายของผู้รับจ้างเอง รวมทั้งผู้รับจ้างจะต้องชดใช้ค่าเสียหายนั้นๆ ตลอดจนค่าใช้จ่ายใดๆ อันเกิดจากการถูกเรียกร้อง หรือถูกฟ้องร้องให้แก่ผู้ว่าจ้างทันที
</P>

  <P class='t-12 tab2'><b>ข้อ ๗.	การตรวจรับงานจ้าง</b></P>
    <P class='t-12 tab3'> เมื่อผู้ว่าจ้างได้ตรวจรับงานจ้างที่ส่งมอบและเห็นว่าถูกต้องครบถ้วนตามบันทึกข้อตกลงแล้ว
ผู้ว่าจ้างจะออกหลักฐานการรับมอบเป็นหนังสือไว้ให้เพื่อผู้รับจ้างนำมาเป็นหลักฐานประกอบการขอรับเงินค่างานจ้างนั้น</P>
  <P class='t-12 tab3'>ถ้าผลของการตรวจรับงานจ้างปรากฏว่างานจ้างที่ผู้รับจ้างส่งมอบไม่ตรงตามบันทึกข้อตกลงผู้ว่าจ้างทรง 
ไว้ซึ่งสิทธิที่จะไม่รับงานจ้างนั้นในกรณีเช่นว่านี้ ผู้รับจ้างต้องทำการแก้ไขให้ถูกต้องตามบันทึกข้อตกลงด้วยค่าใช้จ่ายของผู้รับจ้างเอง
และระยะเวลาที่เสียไปเพราะเหตุดังกล่าวผู้รับจ้างจะนำมาอ้างเป็นเหตุขอขยายเวลาส่งมอบงานจ้างตามบันทึกข้อตกลงหรือของด หรือลดค่าปรับไม่ได้</P>

   <P class='t-12 tab2'><b>ข้อ ๘.	รายละเอียดของงานจ้างคลาดเคลื่อน</b></P>
    <P class='t-12 tab3'>ผู้รับจ้างรับรองว่าได้ตรวจสอบและทำความเข้าใจในรายละเอียดของงานจ้างโดยถี่ถ้วนแล้ว 
หากปรากฏว่ารายละเอียดของงานจ้างนั้นผิดพลาดหรือคลาดเคลื่อนไปจากหลักการทางวิศวกรรมหรือทางเทคนิค 
ผู้รับจ้างตกลงที่จะปฏิบัติตามคำวินิจฉัยของผู้ว่าจ้าง คณะกรรมการตรวจรับพัสดุ เพื่อให้งานแล้วเสร็จบริบูรณ์ คำวินิจฉัยดังกล่าวให้ถือเป็นที่สุด
โดยผู้รับจ้างจะคิดค่าจ้าง ค่าเสียหาย หรือค่าใช้จ่ายใดๆ เพิ่มขึ้นจากผู้ว่าจ้าง หรือขอขยายอายุบันทึกข้อตกลงไม่ได้</P>
  
  <P class='t-12 tab2'><b>ข้อ ๙.	ค่าปรับ</b></P>
    <P class='t-12 tab3'>หากผู้รับจ้างไม่สามารถทำงานให้แล้วเสร็จภายในเวลาที่กำหนดไว้ในบันทึกข้อตกลง
และผู้ว่าจ้างยังมิได้บอกเลิกบันทึกข้อตกลง ผู้รับจ้างจะต้องชำระค่าปรับให้แก่ผู้ว่าจ้างเป็นจำนวนเงิน
วันละ {dataResult.DailyFineRate} บาท ( {CommonDAO.NumberToThaiText(dataResult.DailyFineRate ?? 0)} ) นับถัดจากวันที่ครบกำหนดเวลาแล้วเสร็จของงานตามบันทึกข้อตกลง 
หรือวันที่ผู้ว่าจ้างได้ขยายเวลาทำงานให้จนถึงวันที่ทำงานแล้วเสร็จจริง นอกจากนี้ ผู้รับจ้างยอมให้ผู้ว่าจ้างเรียกค่าเสียหาย 
อันเกิดขึ้นจากการที่ผู้รับจ้างทำงานล่าช้าเฉพาะส่วนที่เกินกว่าจำนวนค่าปรับดังกล่าวได้อีกด้วย</P>
     <P class='t-12 tab3'>ในระหว่างที่ผู้ว่าจ้างยังมิได้บอกเลิกบันทึกข้อตกลงนั้น หากผู้ว่าจ้างเห็นว่าผู้รับจ้างจะไม่สามารถ 
ปฏิบัติตามบันทึกข้อตกลงต่อไปได้ ผู้ว่าจ้างจะใช้สิทธิบอกเลิกบันทึกข้อตกลงและใช้สิทธิตามข้อ ๑๐ ก็ได้และถ้าผู้ว่าจ้าง 
ได้แจ้งข้อเรียกร้องไปยังผู้รับจ้างเมื่อครบกำหนดเวลาแล้วเสร็จของงานขอให้ชำระค่าปรับแล้วผู้ว่าจ้างมีสิทธิที่จะปรับผู้รับจ้างจน ถึงวันบอกเลิกบันทึกข้อตกลงได้อีกด้วย</P>

<P class='t-12 tab2'><b>ข้อ ๑๐.สิทธิของผู้ว่าจ้างภายหลังบอกเลิกบันทึกข้อตกลง
</b></P>
    <P class='t-12 tab3'>ในกรณีที่ผู้ว่าจ้างบอกเลิกบันทึกข้อตกลง ผู้ว่าจ้างอาจทำงานนั้นเองหรือว่าจ้างผู้อื่นให้ทำงานนั้นต่อจน  
แล้วเสร็จก็ได้และในกรณีดังกล่าว ผู้รับจ้างจะต้องรับผิดชอบในค่าเสียหายที่เกิดขึ้นรวมทั้งค่าใช้จ่ายที่เพิ่มขึ้นในการทำงานนั้นต่อ 
ให้แล้วเสร็จตามบันทึกข้อตกลง และผู้ว่าจ้างจะหักเอาจากจำนวนเงินใดๆ ที่จะจ่ายให้แก่ผู้รับจ้างก็ได้</P>

      <P class='t-12 tab2'><b>ข้อ ๑๑. อื่นๆ </b></P>
     <P class='t-12 tab3'>การจ้างเหมาบริการตามบันทึกข้อตกลงนี้ ไม่ทำให้ผู้รับจ้างมีฐานะเป็นลูกจ้างของผู้ว่าจ้างหรือมีความสัมพันธ์ 
ในฐานะเป็นลูกจ้างตามกฎหมายแรงงาน หรือกฎหมายว่าด้วยประกันสังคม
</P>
    <P class='t-12 tab3'>บันทึกข้อตกลงนี้ทำขึ้นเป็นบันทึกข้อตกลงอิเล็กทรอนิกส์ คู่ตกลงได้อ่าน เข้าใจเงื่อนไข และยอมรับเงื่อนไข และได้ยืนยันว่าเป็นผู้มีอำนาจลงนามในบันทึกข้อตกลง จึงได้ลงลายมืออิเล็กทรอนิกส์พร้อมทั้งประทับตรา (ถ้ามี) ในบันทึกข้อตกลงไว้ และต่างฝ่ายต่างยึดถือไว้ฝ่ายละหนึ่งฉบับในระบบของตน </P>

  
</br>
</br>
<!-- 🔹 รายชื่อผู้ลงนาม -->
{signatoryTableHtml}
    <P class='t-12 tab2'> ข้าพเจ้าขอรับรองว่า ทั้งสองฝ่ายได้ลงนามในบันทึกข้อตกลงโดยวิธีการอิเล็กทรอนิกส์ เพื่อแสดงเจตนาของคู่ตกลงแล้ว ข้าพเจ้าจึงได้ลงลายมือชื่ออิเล็กทรอนิกส์รับรองเป็นพยานในบันทึกข้อตกลงพร้อมนี้ </P>

{signatoryTableHtmlWitnesses}
</div>
</body>
</html>
";


            return html;
        }
        catch(Exception ex) 
        { 
        return null;
        }
 
    }
}
