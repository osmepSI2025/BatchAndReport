using BatchAndReport.DAO;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Signatures;
using System.Globalization;
using System.Text;
public class WordEContract_SupportSMEsService
{
    private readonly WordServiceSetting _w;
    private readonly Econtract_Report_GADAO _e;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    private readonly E_ContractReportDAO _eContractReportDAO;
    public WordEContract_SupportSMEsService(WordServiceSetting ws, Econtract_Report_GADAO e
        , IConverter pdfConverter, E_ContractReportDAO eContractReportDAO
        )
    {
        _w = ws;
        _e = e;
        _pdfConverter = pdfConverter; // กำหนดค่า DI สำหรับ PDF Converter
        _eContractReportDAO = eContractReportDAO; // กำหนดค่า DI สำหรับ E_ContractReportDAO
    }


    public async Task<string> OnGetWordContact_SupportSMEsService_HtmlToPDF(string id,string typeContact)
    {
        var result = await _e.GetGAAsync(id);
   
        if (result == null)
            throw new Exception("ไม่พบข้อมูลสัญญารับเงินอุดหนุนสำหรับ SMEs ที่ระบุ");

        // อ่านไฟล์โลโก้และแปลงเป็น Base64
        var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
        string logoBase64 = "";
        if (System.IO.File.Exists(logoPath))
        {
            var bytes = System.IO.File.ReadAllBytes(logoPath);
            logoBase64 = Convert.ToBase64String(bytes);
        }
        // Read CSS file content
        var cssPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "css", "contract.css");
        string contractCss = "";
        if (File.Exists(cssPath))
        {
            contractCss = File.ReadAllText(cssPath, Encoding.UTF8);
        }
        // สร้าง path ฟอนต์แบบ absolute
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabun.ttf");
        string fontBase64 = "";
        if (File.Exists(fontPath))
        {
            var bytes = File.ReadAllBytes(fontPath);
            fontBase64 = Convert.ToBase64String(bytes);
        }

        string signDate = CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now);
        string stringGrantAmount = CommonDAO.NumberToThaiText(result.GrantAmount ?? 0);
        string stringGrantStartDate = CommonDAO.ToThaiDateStringCovert(result.GrantStartDate ?? DateTime.Now);
        string stringGrantEndDate = CommonDAO.ToThaiDateStringCovert(result.GrantEndDate ?? DateTime.Now);

        var findYear = CommonDAO.ToThaiDateString(result.ContractSignDate ?? DateTime.Now);
        string yearThai = findYear[2]; // ดึงปีไทย 4 หลักล่าสุด

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
        // ► Fallback: ถ้าจบลูปแล้วยังไม่มีตราประทับ แต่คุณ “ต้องการให้มีอย่างน้อย placeholder 1 ครั้ง”
        //if (!sealAdded)
        //{

        //    sealAdded = true;
        //}

        //// ► ประกอบผลลัพธ์
        //var signatoryWithLogoHtml = new StringBuilder();
        //if (companySealHtml.Length > 0) signatoryWithLogoHtml.Append(companySealHtml);
        //signatoryWithLogoHtml.Append(signatoryHtml);


        // สร้าง HTML สำหรับสัญญา 


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

    <div class='text-center'>
         <img src='data:image/jpeg;base64,{logoBase64}'  height='80' />
    </div>
</br>
</br>
    <div class='t-14 text-center'><B>สัญญารับเงินอุดหนุน</B></div>
    <div class='t-14 text-center'><B>เพื่อสนับสนุนและยกระดับศักยภาพผู้ประกอบการวิสาหกิจขนาดกลางและขนาดย่อม</B></div>
    <div class='t-14 text-center'><B>ผ่านระบบผู้ให้บริการทางธุรกิจ ปี {yearThai}</B></div>
</br>
    <div class=' t-12 text-right'>ทะเบียนผู้รับเงินอุดหนุนเลขที่ {CommonDAO.ConvertStringArabicToThaiNumerals(result.TaxID)}</div>
    <div class=' t-12 text-right'>เลขที่สัญญา {CommonDAO.ConvertStringArabicToThaiNumerals(result.Contract_Number)}</div>
</br>
    <p class='t-12 tab2'>สัญญาฉบับนี้ทำขึ้น ณ {CommonDAO.ConvertStringArabicToThaiNumerals(result.SignAddress)} เมื่อ {signDate} ระหว่าง</P>
    <p class='t-12 tab2'><B>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B> โดย {CommonDAO.ConvertStringArabicToThaiNumerals(result.SignatoryName)} ผู้มีอำนาจกระทำการแทนสำนักงานฯ 
ซึ่งต่อไปในสัญญานี้เรียกว่า “ผู้ให้เงินอุดหนุน” ฝ่ายหนึ่ง กับ</P>
    <p class='t-12 tab2'><B>ผู้ประกอบการวิสาหกิจขนาดกลางและขนาดย่อม</B> ราย {CommonDAO.ConvertStringArabicToThaiNumerals(result.ContractPartyName)} ซึ่งจดทะเบียนเป็น {CommonDAO.ConvertStringArabicToThaiNumerals(result.RegType)} 
เลขประจำตัวผู้เสียภาษี {CommonDAO.ConvertStringArabicToThaiNumerals(result.TaxID)} 
        ณ {signDate}  มีสำนักงานใหญ่
        ตั้งอยู่เลขที่ {CommonDAO.ConvertStringArabicToThaiNumerals(result.HQLocationAddressNo)}
        ตำบล/แขวง {CommonDAO.ConvertStringArabicToThaiNumerals(result.HQLocationDistrict)} อำเภอ/เขต {CommonDAO.ConvertStringArabicToThaiNumerals(result.HQLocationDistrict)} จังหวัด {result.HQLocationProvince}  {CommonDAO.ConvertStringArabicToThaiNumerals(result.HQLocationZipCode)}
        ไปรษณีย์อิเล็กทรอนิกส์(E-mail) {CommonDAO.ConvertStringArabicToThaiNumerals(result.RegEmail)} โดย {CommonDAO.ConvertStringArabicToThaiNumerals(result.RegPersonalName)} บัตรประจำตัวประชาชนเลขที่ {CommonDAO.ConvertStringArabicToThaiNumerals(result.RegIdenID)}
        ผู้มีอำนาจลงนามผูกพัน {CommonDAO.ConvertStringArabicToThaiNumerals(result.RegType)} ปรากฏตามสำเนา
        หนังสือรับรอง {CommonDAO.ConvertStringArabicToThaiNumerals(result.RegType)} ของสำนักงานทะเบียน 
        หุ้นส่วนบริษัท {CommonDAO.ConvertStringArabicToThaiNumerals(result.ContractPartyName)} ลง {signDate})
        ซึ่งต่อไปในสัญญานี้ เรียกว่า “ผู้รับเงินอุดหนุน” อีกฝ่ายหนึ่ง
    </P>
    <p class='t-12 tab2'>ทั้งสองฝ่ายได้ตกลงทำสัญญากัน มีข้อความดังต่อไปนี้</P>
    <p class='t-12 tab2'>ข้อ ๑ ผู้ให้เงินอุดหนุนตกลงให้เงินอุดหนุนและผู้รับเงินอุดหนุนตกลงรับเงินอุดหนุน  
จำนวน {CommonDAO.ConvertStringArabicToThaiNumerals((result.GrantAmount ?? 0).ToString("N0"))} บาท ({stringGrantAmount})
ตั้งแต่ {stringGrantStartDate} ถึง {stringGrantEndDate}
โดยให้ผู้รับการอุดหนุนเข้ารับการพัฒนา เพื่อใช้จ่ายในการ {CommonDAO.ConvertStringArabicToThaiNumerals(result.SpendingPurpose)}
จากการให้ความช่วยเหลือ อุดหนุน จากเงินกองทุนส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม
        ผ่านผู้ให้บริการ ทางธุรกิจ ปี ๒๕๖๗ ภายใต้โครงการส่งเสริมผู้ประกอบการผ่านระบบ BDS ระยะเวลาดำเนินการ ๒ ปี 
(ปี ๒๕๖๗-๒๕๖๘) ตามข้อเสนอการพัฒนาซึ่งได้รับอนุมัติจากผู้ให้เงินอุดหนุน ตามระเบียบคณะกรรมการ
บริหารสำนักงานส่งเสริม วิสาหกิจขนาดกลางและขนาดย่อม ว่าด้วยหลักเกณฑ์ เงื่อนไข และวิธีการให้ความ
ช่วยเหลือ อุดหนุน วิสาหกิจ ขนาดกลางและขนาดย่อม จากเงินกองทุนส่งเสริมวิสาหกิจขนาดกลางและ
ขนาดย่อม พ.ศ. ๒๕๖๔ ประกาศ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เรื่อง เชิญชวน
หน่วยงานที่ประสงค์ขึ้นทะเบียนผู้ให้บริการ ทางธุรกิจ เพื่อสนับสนุนและยกระดับศักยภาพผู้ประกอบการ
วิสาหกิจขนาดกลางและขนาดย่อม และเชิญชวน วิสาหกิจขนาดกลางและขนาดย่อม ยื่นความประสงค์ขอรับ
ความช่วยเหลือ อุดหนุนจากเงินกองทุนส่งเสริม วิสาหกิจขนาดกลางและขนาดย่อม ผ่านผู้ให้บริการทางธุรกิจ 
ปี ๒๕๖๗ และประกาศสำนักงานส่งเสริมวิสาหกิจ ขนาดกลางและขนาดย่อม เรื่อง เชิญชวนหน่วยงานที่
ประสงค์ขึ้นทะเบียนผู้ให้บริการทางธุรกิจ เพื่อสนับสนุน และยกระดับศักยภาพผู้ประกอบการวิสาหกิจขนาด
กลางและขนาดย่อมและเชิญชวนวิสาหกิจขนาดกลางและ ขนาดย่อม ยื่นความประสงค์ขอรับความช่วยเหลือ 
อุดหนุนฯ (ฉบับที่ ๒) และผู้รับเงินอุดหนุนต้องดำเนิน กิจกรรมและใช้จ่ายเงินตามแผนการดำเนินงานและ
แผนการใช้จ่ายที่ระบุไว้ในข้อเสนอการพัฒนาที่ได้รับอนุมัติ อย่างเคร่งครัด และให้ถือว่าเป็นส่วนหนึ่งของ
สัญญาฉบับนี้
    </P>
    <p class='t-12 tab2'>ข้อ ๒ ผู้รับเงินอุดหนุนจะต้องสำรองเงินจ่ายไปก่อน แล้วจึงนำต้นฉบับใบเสร็จรับเงินมาเบิก
กับ ผู้ให้เงินอุดหนุน วงเงินไม่เกินตามข้อ ๑ ทั้งนี้ ผู้ให้เงินอุดหนุนจะสนับสนุนจำนวนเงินตามจำนวนที่
จ่ายจริงและ เป็นไปตามสัดส่วนการร่วมค่าใช้จ่ายในการสนับสนุนระหว่างผู้ให้เงินอุดหนุนและผู้รับเงินอุดหนุน โดยสัดส่วน งบประมาณที่ให้การอุดหนุนดังกล่าวต้องเป็นไปตามการจัดกลุ่มและสัดส่วนของผู้ประกอบการ ตามประกาศ แนบท้ายสัญญา</P>
   
<p class='t-12 tab2'>ในการให้ความช่วยเหลือ อุดหนุน วิสาหกิจขนาดกลางและขนาดย่อม ผ่านผู้ให้บริการทาง
ธุรกิจ ผู้รับเงินอุดหนุนจะได้รับความช่วยเหลือ อุดหนุน ในโครงการนี้ หรือโครงการให้ความช่วยเหลือ อุดหนุน ผ่านผู้ให้บริการทางธุรกิจในปีอื่นๆ ในวงเงินรวมกันสูงสุดไม่เกิน ๕๐๐,๐๐๐ บาท (ห้าแสนบาทถ้วน) ตลอดระยะ
เวลา การดำเนินธุรกิจ  ดังนั้น วงเงินที่ได้รับการอุดหนุนตามสัญญานี้ จะต้องถูกหักจากวงเงินรวมที่ได้รับสิทธิ์ <br></P>
  
<p class='t-12 tab2'>ข้อ ๓ เมื่อผู้รับเงินอุดหนุนดำเนินกิจกรรมเข้ารับการพัฒนาเสร็จสมบูรณ์แล้วตามแผนการ
ดำเนิน กิจกรรมในข้อเสนอการพัฒนา และนำส่งรายงานผลการพัฒนาและรายละเอียดที่เกี่ยวข้องมายังผู้
ให้เงิน อุดหนุน โดยผู้รับเงินอุดหนุนต้องเบิกค่าใช้จ่ายทันทีหลังจากได้รับการพัฒนาหรือก่อนสิ้นสุดสัญญา
ฉบับนี้ ภายใน ๓๐ (สามสิบ) วันทำการ นับจากวันที่สิ้นสุดสัญญา</P>
    <p class='t-12 tab2'>ข้อ ๔ ผู้รับเงินอุดหนุนยินยอมรับผิดชอบค่าใช้จ่ายส่วนเกินจากการสนับสนุนตามการให้ความ ช่วยเหลือในโครงการนี้ที่ได้กำหนดไว้ รวมทั้งรับผิดชอบภาษีมูลค่าเพิ่ม และภาษีอื่น ๆ (ถ้ามี) ที่เกิดจาก ค่าใช้จ่ายที่ขอรับการอุดหนุน</P>
    <p class='t-12 tab2'>ข้อ ๕ เงินที่ผู้รับเงินอุดหนุนได้รับจากโครงการนี้ เป็นเงินที่รวมภาษี และค่าธรรมเนียมต่างๆ ไว้ทั้งหมดแล้ว และถือเป็นรายได้ของวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งจะต้องถูกหักภาษี ณ ที่จ่าย และ ต้องเสียภาษีตามที่กฎหมายกำหนด และหากวิสาหกิจขนาดกลางและขนาดย่อมเป็นผู้ซึ่งจดทะเบียน ภาษีมูลค่าเพิ่ม จะต้องมีการแสดงรายการคำนวณภาษีมูลค่าเพิ่มไว้ให้ชัดเจนปรากฏไว้ในใบสำคัญการรับเงิน หรือใบเสร็จรับเงิน หรือใบกำกับภาษี ที่ยื่นให้ผู้ให้เงินอุดหนุน โดยวิสาหกิจขนาดกลางและขนาดย่อมมีหน้าที่ จะต้องนำเงินที่ได้รับดังกล่าว ไปประกอบการคำนวณรายได้เพื่อเสียภาษีเงินได้ในปีที่เกิดรายได้ด้วย </P>
    <p class='t-12 tab2'>ข้อ ๖ กรณีการโอนเงินให้แก่ผู้รับเงินอุดหนุน ผู้ให้เงินอุดหนุนจะใช้วิธีการโอนเงินผ่านระบบ อิเล็กทรอนิกส์ และหากมีค่าธรรมเนียมการโอนเงิน ผู้รับเงินอุดหนุนจะเป็นผู้รับผิดชอบค่าธรรมเนียมในการ โอนเงินดังกล่าว</P>
    <p class='t-12 tab2'>ข้อ ๗ ผู้รับเงินอุดหนุนจะเปลี่ยนแปลงข้อเสนอการพัฒนาและวงเงินอุดหนุนตามที่ได้รับ
อนุมัติ จากผู้ให้เงินอุดหนุนได้ ต่อเมื่อผู้รับเงินอุดหนุนได้แจ้งเป็นหนังสือให้ผู้ให้เงินอุดหนุนทราบ และ
ได้รับความ เห็นชอบเป็นหนังสือจากผู้ให้เงินอุดหนุนก่อนทุกครั้ง โดยผู้รับเงินอุดหนุนจะต้องดำเนินการ
ก่อนวันสิ้นสุด สัญญาไม่น้อยกว่า ๓๐ (สามสิบ) วันทำการ</P>
    <p class='t-12 tab2'>ข้อ ๘ ผู้รับเงินอุดหนุนจะต้องใช้จ่ายเงินอุดหนุนเพื่อดำเนินการตามข้อเสนอการพัฒนา ซึ่งได้รับการอนุมัติ ให้เป็นไปตามวัตถุประสงค์และกิจกรรมตามข้อเสนอการพัฒนาเท่านั้น โดยผู้รับเงินอุดหนุน ตกลงยินยอมให้ผู้ให้เงินอุดหนุนตรวจสอบผลการปฏิบัติงาน และการใช้จ่ายเงินอุดหนุนที่ได้รับ และผู้รับเงิน อุดหนุนมีหน้าที่ต้องรายงานผลการปฏิบัติงานและการใช้จ่ายเงินอุดหนุนที่รับตามแบบและภายในเวลาที่ กำหนด </P>

    <p class='t-12 tab2'>ข้อ ๙ กรณีที่มีการตรวจพบในภายหลังว่าผู้รับเงินอุดหนุนขาดคุณสมบัติในการรับเงินอุดหนุน 
ผู้ให้เงินอุดหนุนมีสิทธิบอกเลิกสัญญาได้ทันที หรือในกรณีผู้รับเงินอุดหนุนนำเงินไปใช้ผิดจากวัตถุประสงค์ตาม 
ข้อเสนอการพัฒนา ผู้รับเงินอุดหนุนจะต้องรับผิดชอบชดใช้เงินอุดหนุนที่ได้รับไปทั้งหมดคืนให้แก่ผู้ให้เงินอุด
หนุน ภายใน ๓๐ (สามสิบ) วัน นับแต่วันที่ได้รับหนังสือแจ้งจากผู้ให้เงินอุดหนุน พร้อมด้วยดอกเบี้ยในอัตรา
ร้อยละ ๕ (ห้า) ต่อปี นับแต่วันที่ได้รับเงินอุดหนุนจนกว่าจะชดใช้เงินคืนจนครบถ้วนเสร็จสิ้น </P>

    <p class='t-12 tab2'>ข้อ ๑๐ ในกรณีผู้รับเงินอุดหนุนไม่ปฏิบัติตามสัญญาข้อหนึ่งข้อใด ผู้ให้เงินอุดหนุนจะมีหนัง
สือแจ้ง ให้ผู้รับเงินอุดหนุนทราบ โดยจะกำหนดระยะเวลาพอสมควรเพื่อให้ปฏิบัติให้ถูกต้องตามสัญญา 
และหาก ผู้รับเงินอุดหนุนไม่ปฏิบัติภายในระยะเวลาที่กำหนดดังกล่าว ผู้ให้เงินอุดหนุนมีสิทธิบอกเลิกสัญญา
ได้ทันที โดย มีหนังสือบอกเลิกสัญญาแจ้งให้ผู้รับเงินอุดหนุนทราบ</P>
    <p class='t-12 tab2'>ข้อ ๑๑ ในกรณีที่มีการบอกเลิกสัญญาตามข้อ ๑๐ ผู้รับเงินอุดหนุนจะต้องชดใช้เงินอุดหนุน
คืน ให้แก่ผู้ให้เงินอุดหนุนตามจำนวนเงินที่ได้รับทั้งหมด หรือตามจำนวนเงินคงเหลือในวันบอกเลิกสัญญา 
หรือตาม จำนวนเงินที่ผู้ให้เงินอุดหนุนจะพิจารณาตามความเหมาะสมแล้วแต่กรณี ซึ่งผู้ให้เงินอุดหนุนจะแจ้ง
เป็นหนังสือ พร้อมการบอกเลิกสัญญา ให้ผู้รับเงินอุดหนุนทราบว่าต้องชดใช้เงินคืนจำนวนเท่าใด โดยผู้รับเงิน
อุดหนุนต้อง ชำระเงินดังกล่าวพร้อมด้วยดอกเบี้ยในอัตรา ร้อยละ ๕ (ห้า) ต่อปี นับแต่วันบอกเลิกสัญญา
จนถึงวันที่ชดใช้ เงินคืนจนครบถ้วนเสร็จสิ้น ทั้งนี้ ในกรณีเกิดความเสียหายอย่างหนึ่งอย่างใดแก่ผู้ให้เงิน
อุดหนุน ผู้ให้เงิน อุดหนุนมีสิทธิที่จะเรียกค่าเสียหายจากผู้รับเงินอุดหนุนอีกด้วย</P>
    <p class='t-12 tab2'>ข้อ ๑๒ ผู้รับเงินอุดหนุนต้องปฏิบัติตามเงื่อนไขที่กำหนดไว้ในระเบียบและประกาศแนบท้าย สัญญานี้</P>

    <p class='t-12 tab2'>ข้อ ๑๓ ที่อยู่ของผู้รับเงินอุดหนุนที่ปรากฏในสัญญานี้ ให้ถือว่าเป็นภูมิลำเนาของผู้รับเงิน
อุดหนุน การส่งหนังสือ คำบอกกล่าวทวงถาม จดหมาย หรือเอกสารอื่นใดไปยังผู้รับเงินอุดหนุน ให้ส่ง
ไปยังภูมิลำเนา ผู้รับเงินอุดหนุนดังกล่าว และให้ถือว่าเป็นการส่งโดยชอบ โดยถือว่าผู้รับเงินอุดหนุน
ได้ทราบข้อความ ในเอกสารดังกล่าวนับแต่วันที่หนังสือ คำบอกกล่าวทวงถาม จดหมาย หรือเอกสารอื่นใด
ไปถึงภูมิลำเนา ของผู้รับเงินอุดหนุน ไม่ว่าผู้รับเงินอุดหนุนหรือบุคคลอื่นใดที่พักอาศัยอยู่ในภูมิลำเนาของผู้
รับเงินอุดหนุนจะ ได้รับหนังสือ คำบอกกล่าวทวงถาม จดหมาย หรือเอกสารนั้นไว้หรือไม่ก็ตาม</P>
    <p class='t-12 tab2'>ถ้าผู้รับเงินอุดหนุนเปลี่ยนแปลงสถานที่อยู่ หรือไปรษณีย์อิเล็กทรอนิกส์ (E-mail) ผู้รับเงิน
อุดหนุน มีหน้าที่แจ้งให้ผู้ให้เงินอุดหนุนทราบภายใน ๗ (เจ็ด) วัน นับแต่วันเปลี่ยนแปลงสถานที่อยู่หรือ
ไปรษณีย์อิเล็กทรอนิกส์ (E-mail) หากผู้รับเงินอุดหนุนไม่แจ้งการเปลี่ยนแปลงสถานที่อยู่และผู้ให้เงิน
อุดหนุนได้ส่ง หนังสือ คำบอกกล่าวทวงถาม จดหมาย หรือเอกสารอื่นใดไปยังผู้รับเงินอุดหนุนตามที่อยู่ที่
ปรากฏในสัญญานี้ ให้ถือว่าผู้รับเงินอุดหนุนได้ทราบข้อความในเอกสารดังกล่าวโดยชอบตามวรรคหนึ่งแล้ว</P>
    <p class='t-12 tab2'>สัญญานี้ทำขึ้นเป็นสัญญาอิเล็กทรอนิกส์ คู่สัญญาได้อ่าน เข้าใจเงื่อนไข และยอมรับเงื่อนไข และได้ยืนยันว่าเป็นผู้มีอำนาจลงนามในสัญญาจึงได้ลงลายมืออิเล็กทรอนิกส์พร้อมทั้งประทับตรา (ถ้ามี) ในสัญญาไว้ และต่างฝ่ายต่างยึดถือไว้ฝ่ายละหนึ่งฉบับในระบบของตน</P>

</br>
</br>
<div class='t-12'>
{signatoryTableHtml}
    <P class='t-12 tab2'>ข้าพเจ้าขอรับรองว่า ทั้งสองฝ่ายได้ลงนามในสัญญาโดยวิธีการอิเล็กทรอนิกส์ เพื่อแสดงเจตนาของคู่สัญญาแล้ว ข้าพเจ้าจึงได้ลงลายมือชื่ออิเล็กทรอนิกส์รับรองเป็นพยานในสัญญาพร้อมนี้
</P>

{signatoryTableHtmlWitnesses}
</div>
</body>
</html>
";
     
        return html;
    }
}
