using BatchAndReport.DAO;
using BatchAndReport.Entities;
using BatchAndReport.Models;
using BatchAndReport.Services;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using Spire.Doc.Documents;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;


public class WordEContract_HireEmployee
{
    private readonly WordServiceSetting _w;
    private readonly Econtract_Report_ECDAO _e;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    private readonly EContractDAO _eContractDAO;
    private readonly E_ContractReportDAO _eContractReportDAO;

    public WordEContract_HireEmployee(WordServiceSetting ws, Econtract_Report_ECDAO e
         , IConverter pdfConverter
        ,
EContractDAO eContractDAO
        , E_ContractReportDAO eContractReportDAO
        )
    {
        _w = ws;
        _e = e;
        _pdfConverter = pdfConverter;
        _eContractDAO = eContractDAO;
        _eContractReportDAO = eContractReportDAO;
    }
    #region   4.1.3.3. สัญญาจ้างลูกจ้าง

    public async Task<string> OnGetWordContact_HireEmployee_ToPDF(string id, string typeContact)
    {
        try
        {
            // ── 0) validate args / DI ───────────────────────────────────────────────
            if (string.IsNullOrWhiteSpace(id))
                throw new ArgumentException("id is required.", nameof(id));
            if (string.IsNullOrWhiteSpace(typeContact))
                throw new ArgumentException("typeContact is required.", nameof(typeContact));

            if (_e == null) throw new NullReferenceException("_e is null");
            if (_eContractReportDAO == null) throw new NullReferenceException("_eContractReportDAO is null");
            // if (_pdfConverter == null)       throw new NullReferenceException("_pdfConverter is null"); // ถ้าใช้ convert จริงค่อยเปิด

            // ── 1) โหลดข้อมูลหลัก (กัน result = null) ─────────────────────────────
            var result = await _e.GetECAsync(id);
            if (result == null)
                throw new InvalidOperationException($"No data found for id '{id}'.");

            // ── 2) path ต่าง ๆ ─────────────────────────────────────────────────────
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
            // ── 3) ข้อความ/วันที่ (ป้องกัน null ด้วย ??) ────────────────────────────
            string strAttorneyLetterDate = CommonDAO.ToThaiDateStringCovert(result.AttorneyLetterDate ?? DateTime.Now);
            string strAttorney =
                result.AttorneyFlag == true
                ? $"ผู้รับมอบหมายตามคำสั่งสำนักงานฯ ที่ {result.AttorneyLetterNumber ?? ""} ลง {strAttorneyLetterDate}"
                : "";

            string strcontractsign = CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now);
            string strHiringStart = CommonDAO.ToThaiDateStringCovert(result.HiringStartDate ?? DateTime.Now);
            string strHiringEnd = CommonDAO.ToThaiDateStringCovert(result.HiringEndDate ?? DateTime.Now);
            string strSalary = CommonDAO.NumberToThaiText(result.Salary ?? 0);

            #region signlist joa
            // call function RenderSignatory
            var signatoryTableHtml = "";

            if (result.Signatories.Count > 0)
            {
                signatoryTableHtml = await _eContractReportDAO.RenderSignatory_2Column(result.Signatories);


            }

            var signatoryTableHtmlWitnesses = "";

            if (result.Signatories.Count > 0)
            {
                signatoryTableHtmlWitnesses = await _eContractReportDAO.RenderSignatory_Witnesses_2Column(result.Signatories);
            }


            #endregion signlist

            #region cleanCode
            // ตัวอย่างการใช้ Regex เพื่อลบ style attribute ออก
            var cleanDescription = Regex.Replace(CommonDAO.ConvertStringArabicToThaiNumerals(result.Work_Detail), "style=\"[^\"]*\"", string.Empty);

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
            // ⭐ เพิ่ม class="t-12" ให้กับทุก <p>
            foreach (var p in htmlDoc.DocumentNode.Descendants("p"))
            {
                var existingClass = p.GetAttributeValue("class", "");
                if (!existingClass.Contains("t-12"))
                {
                    p.SetAttributeValue("class", (existingClass + " t-12 tab3").Trim());
                }
            }

            string cleanedHtml = htmlDoc.DocumentNode.OuterHtml;

            #endregion


            // ── 5) เนื้อหา HTML (ใช้ ?? "" กัน null string) ────────────────────────
            string htmlBody = $@"
<div style='margin-bottom:24px;text-align:center;'>
    {(System.IO.File.Exists(logoPath) ? $" <img src='data:image/jpeg;base64,{logoBase64}' height='80' />" : "")}
</div>
<div class='text-center t-14'><b>สัญญาจ้างลูกจ้าง</b></div>
</br>
<p class='tab2 t-12'>
   สัญญาฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เลขที่ ๑๒๐ หมู่ ๓ 
ศูนย์ราชการเฉลิมพระเกียรติ ๘๐ พรรษา ๕ ธันวาคม ๒๕๕๐ (อาคารซี) ชั้น ๒, ๑๐, ๑๑ ถนนแจ้งวัฒนะ แขวงทุ่งสองห้อง เขตหลักสี่ กรุงเทพ ๑๐๒๑๐
 เมื่อ {strcontractsign}
</p>
<p class='tab2 t-12'>
    ระหว่าง สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม โดย {CommonDAO.ConvertStringArabicToThaiNumerals(result.OSMEP_NAME) ?? ""} ตำแหน่ง {CommonDAO.ConvertStringArabicToThaiNumerals(result.OSMEP_POSITION) ?? ""} {CommonDAO.ConvertStringArabicToThaiNumerals(strAttorney)} ซึ่งต่อไปในสัญญานี้จะเรียกว่า <b>“ผู้ว่าจ้าง”</b>
</p>
<p class='tab2 t-12'>
    ฝ่ายหนึ่ง กับ {result.EmploymentName ?? ""} เลขประจำตัวประชาชน {CommonDAO.ConvertStringArabicToThaiNumerals(result.IdenID) ?? ""} อยู่บ้านเลขที่ {CommonDAO.ConvertStringArabicToThaiNumerals(result.EmpAddress) ?? ""} ซึ่งต่อไปในสัญญานี้จะเรียกว่า <b>“ลูกจ้าง”</b> 
อีกฝ่ายหนึ่ง
</p>
<p class='tab2 t-12'>
โดยทั้งสองฝ่ายได้ตกลงทำร่วมกันดังมีรายละเอียดต่อไปนี้
</p>
<p class='tab2 t-12'>
    ๑. ผู้ว่าจ้างตกลงจ้างลูกจ้างปฏิบัติงานกับผู้ว่าจ้าง โดยให้ปฏิบัติงานภายใต้โครงการ {CommonDAO.ConvertStringArabicToThaiNumerals(result.Work_Location) ?? ""} ในตำแหน่ง {CommonDAO.ConvertStringArabicToThaiNumerals(result.WorkDetail) ?? ""} โดยมีรายละเอียดหน้าที่ความรับผิดชอบปรากฏตามเอกสารแนบท้ายสัญญาจ้าง ตั้งแต่ {strHiringStart} ถึง {strHiringEnd}
</p>
<p class='tab2 t-12'>
๒. ผู้ว่าจ้างจะจ่ายค่าจ้างให้แก่ลูกจ้างในระหว่างระยะเวลาการปฏิบัติงานของลูกจ้างตามสัญญานี้ ในอัตราเดือนละ {CommonDAO.ConvertCurrencyToThaiNumerals(result.Salary.HasValue ? (int)result.Salary.Value : 0)}  บาท ({strSalary}) 
โดยจะจ่ายให้ในวันทำการก่อนวันทำการสุดท้ายของธนาคารในเดือนนั้นสามวันทำการ และนำเข้าบัญชีเงินฝากของลูกจ้าง ณ ที่ทำการของผู้ว่าจ้าง หรือ ณ ที่อื่นใดตามที่ผู้ว่าจ้างกำหนด 
</p>
 <p class='tab2 t-12'>
  ๓. ในการจ่ายค่าจ้าง และ/หรือ เงินในลักษณะอื่นให้แก่ลูกจ้าง ลูกจ้างตกลงยินยอมให้ผู้ว่าจ้างหักภาษี ณ ที่จ่าย และ/หรือ เงินอื่นใดที่ต้องหักโดยชอบด้วยระเบียบ ข้อบังคับของผู้ว่าจ้างหรือตามกฎหมายที่เกี่ยวข้อง
</p>
<p class='tab2 t-12'>
   ๔. ตลอดระยะเวลาการปฏิบัติงานตามสัญญานี้ ลูกจ้างมีสิทธิได้รับสิทธิประโยชน์อื่น ๆ ตามที่กำหนดไว้ใน ระเบียบ ข้อบังคับ คำสั่ง หรือประกาศใด ๆ ตามที่ผู้ว่าจ้างกำหนด
</p>
<p class='tab2 t-12'>
   ๕. ผู้ว่าจ้างจะทำการประเมินผลการปฏิบัติงานอย่างน้อยปีละสองครั้ง ตามหลักเกณฑ์และวิธีการที่ ผู้ว่าจ้างกำหนด ทั้งนี้ หากผลการประเมินไม่ผ่านตามหลักเกณฑ์ที่กำหนด ผู้ว่าจ้างมีสิทธิบอกเลิกสัญญาจ้างได้ และลูกจ้างไม่มีสิทธิเรียกร้องเงินชดเชยหรือเงินอื่นใด
</p>
<p class='tab2 t-12'>
 ๖. ตลอดระยะเวลาการปฏิบัติงานตามสัญญานี้ ลูกจ้างจะต้องปฏิบัติตามกฎ ระเบียบ ข้อบังคับ คำสั่งหรือประกาศใด ๆ ของผู้ว่าจ้าง ตลอดจนมีหน้าที่ต้องรักษาวินัยและยอมรับการลงโทษทางวินัยของผู้ว่าจ้างโดยเคร่งครัด 
และยินยอมให้ถือว่า กฎหมาย ระเบียบ ข้อบังคับ หรือคำสั่งต่าง ๆ ของผู้ว่าจ้างเป็นส่วนหนึ่งของสัญญาจ้างนี้
</p>
<p class='tab2 t-12'>
 ในกรณีลูกจ้างจงใจขัดคำสั่งโดยชอบของผู้ว่าจ้างหรือละเลยไม่นำพาต่อคำสั่งเช่นว่านั้นเป็นอาจิณ หรือประการอื่นใด อันไม่สมควรกับการปฏิบัติหน้าที่ของตนให้ลุล่วงไปโดยสุจริตและถูกต้อง ลูกจ้างยินยอมให้ผู้ว่าจ้างบอกเลิกสัญญาจ้างโดยมิต้องบอกกล่าวล่วงหน้า
</p>
<p class='tab2 t-12'>
   ๗. ลูกจ้างต้องปฏิบัติงานให้กับผู้ว่าจ้าง ตามที่ได้รับมอบหมายด้วยความซื่อสัตย์ สุจริต และตั้งใจปฏิบัติงานอย่างเต็มกำลังความสามารถของตน โดยแสวงหาความรู้และทักษะเพิ่มเติมหรือกระทำการใด  เพื่อให้ผลงานในหน้าที่มีคุณภาพดีขึ้น 
ทั้งนี้ ต้องรักษาผลประโยชน์และชื่อเสียงของผู้ว่าจ้าง และไม่เปิดเผยความลับหรือข้อมูลของทางราชการให้ผู้หนึ่งผู้ใดทราบ โดยมิได้รับอนุญาตจากผู้รับผิดชอบงานนั้น ๆ 
</p>
<p class='tab2 t-12'>
๘. ลูกจ้างมีหน้าที่ปฏิบัติงานที่เกี่ยวข้องกับการประมวลผลข้อมูลส่วนบุคคลไม่ว่าจะเป็นการเก็บรวบรวม 
การใช้และการเปิดเผยข้อมูลส่วนบุคคลโดยเคร่งครัดตามกฎหมายว่าด้วยการคุ้มครองข้อมูลส่วนบุคคล 
ระเบียบและนโยบายการคุ้มครองข้อมูลส่วนบุคคลของผู้ว่าจ้าง รวมถึงต้องรักษาความลับและความปลอดภัย
ของข้อมูลส่วนบุคคลที่ลูกจ้างได้รับหรือเข้าถึงจากการปฏิบัติงาน ห้ามเปิดเผย ทำสำเนา ส่งต่อ หรือใช้ข้อมูล
ส่วนบุคคลดังกล่าวเพื่อประโยชน์ส่วนตนหรือบุคคลอื่นโดยมิชอบ และต้องแจ้งให้ผู้ว่าจ้างทราบโดยทันที 
หากพบเหตุอันควรสงสัยว่ามีการละเมิดหรือรั่วไหลของข้อมูลส่วนบุคคล ทั้งนี้ 
การฝ่าฝืนหน้าที่ดังกล่าวเป็นเหตุให้เจ้าของข้อมูลส่วนบุคคลหรือผู้ว่าจ้างเสียหาย 
ถือเป็นการผิดสัญญาจ้างงานอย่างร้ายแรงที่ผู้ว่าจ้างมีสิทธิบอกเลิกสัญญาจ้างได้ทันที
</p>
<p class='tab2 t-12'>
    ๙. สัญญานี้สิ้นสุดลงเมื่อเข้ากรณีใดกรณีหนึ่ง ดังต่อไปนี้
</p>
<p class='tab3 t-12'>๙.๑ สิ้นสุดระยะเวลาตามสัญญาจ้าง</p>
<p class='tab3 t-12'>๙.๒ เมื่อผู้ว่าจ้างบอกเลิกสัญญาจ้าง หรือลูกจ้างบอกเลิกสัญญาจ้างตามข้อ ๑๐</p>
<p class='tab3 t-12'>๙.๓ ลูกจ้างกระทำการผิดวินัยร้ายแรง</p>
<p class='tab3 t-12'>๙.๔ ลูกจ้างไม่ผ่านการประเมินผลการปฏิบัติงานของลูกจ้างตามข้อ ๕</p>
<p class='tab2 t-12'>
   ๑๐. ในกรณีที่สัญญาสิ้นสุดตามข้อ ๘ ข้อ ๙.๓ และ ๙.๔ ลูกจ้างยินยอมให้ผู้ว่าจ้างสั่งให้ลูกจ้าง
พ้นสภาพการเป็นลูกจ้างได้ทันที โดยไม่จำเป็นต้องมีหนังสือว่ากล่าวตักเตือน และผู้ว่าจ้างไม่ต้องจ่ายค่าชดเชย
หรือเงินอื่นใดให้แก่ลูกจ้างทั้งสิ้น เว้นแต่ค่าจ้างที่ลูกจ้างจะพึงได้รับตามสิทธิ 

</p>
<p class='tab2 t-12'>
   ๑๑. ลูกจ้างมีสิทธิบอกเลิกสัญญาจ้างได้ก่อนสัญญาครบกำหนด โดยทำหนังสือแจ้งเป็นลายลักษณ์อักษรต่อผู้ว่าจ้างได้ทราบล่วงหน้าไม่น้อยกว่า ๓๐ วัน เมื่อผู้ว่าจ้างได้อนุมัติแล้ว ให้ถือว่าสัญญาจ้างนี้ได้สิ้นสุดลง
</p>
<p class='tab2 t-12'>
  ๑๒. ในกรณีที่ลูกจ้างกระทำการใดอันทำให้ผู้ว่าจ้างได้รับความเสียหาย ไม่ว่าเหตุนั้นผู้ว่าจ้างจะนำมาเป็นเหตุบอกเลิกสัญญาจ้างหรือไม่ก็ตาม ผู้ว่าจ้างมีสิทธิจะเรียกร้องค่าเสียหาย และลูกจ้างยินยอมชดใช้ค่าเสียหายตามที่ผู้ว่าจ้างเรียกร้องทุกประการ 
</p>
<p class='tab2 t-12'>
    ๑๓. ลูกจ้างจะต้องไม่เปิดเผยหรือบอกกล่าวอัตราค่าจ้างของลูกจ้างให้แก่บุคคลใดทราบ ไม่ว่าจะโดยวิธีใดหรือเวลาใด เว้นแต่จะเป็นการกระทำตามกฎหมายหรือคำสั่งศาล
</p>


<p class='tab2 t-12'>
สัญญาฉบับนี้ได้จัดทำขึ้นเป็นสัญญาอิเล็กทรอนิกส์ คู่สัญญาได้อ่าน ตรวจสอบและทำความเข้าใจ ข้อความในสัญญาฉบับนี้โดยละเอียดแล้ว จึงได้ลงลายมือชื่ออิเล็กทรอนิกส์ไว้เป็นหลักฐาน ณ วัน เดือน ปี ดังกล่าวข้างต้น และมีพยานรู้ถึงการลงนามของคู่สัญญา และคู่สัญญาต่างฝ่ายต่างเก็บรักษาไฟล์สัญญาอิเล็กทรอนิกส์ฉบับนี้ไว้เป็นหลักฐาน </p>
</br>
</br>
{signatoryTableHtml} 
</br>
{signatoryTableHtmlWitnesses} 

<div style='page-break-before: always;'></div>
<p class='text-center t-14' style='font-weight:bold;'>เอกสารแนบท้ายสัญญาจ้างลูกจ้าง</p>

</br>

<div class='t-12 editor-content'>
    {cleanedHtml}
</div>


";

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
</head>
<body>
    {htmlBody}
</body>
</html>";

            return html;
        }
        catch (Exception ex) { throw new Exception("Error in OnGetWordContact_HireEmployee: " + ex.Message, ex); }
    }
    #endregion    4.1.3.3. สัญญาจ้างลูกจ้าง
}
