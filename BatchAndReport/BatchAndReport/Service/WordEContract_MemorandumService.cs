using BatchAndReport.DAO;
using BatchAndReport.Models;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Win32;
using Serilog;
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


    #endregion

    #region old
    public async Task<string> OnGetWordContact_MemorandumService_HtmlToPDF(string id, string typeContact)
    {
        var result = await _eContractReportDAO.GetMOUAsync(id);

        if (result == null)
        {
            throw new Exception("ไม่พบข้อมูลบันทึกข้อตกลงความร่วมมือ");
        }

        // Logo
        string strContract_Value = CommonDAO.NumberToThaiText(result.Contract_Value ?? 0);
        string strSign_Date = CommonDAO.ToThaiDateStringCovert(result.Sign_Date ?? DateTime.Now);
        string strStart_Date = CommonDAO.ToThaiDateStringCovert(result.Start_Date ?? DateTime.Now);
        string strEnd_Date = CommonDAO.ToThaiDateStringCovert(result.End_Date ?? DateTime.Now);

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
        string contractLogoHtml = "";
        var logoOrgList = await _eContractReportDAO.Getsp_GetOrganizationLogosAsync(id,"MOU");

        if (logoOrgList.Count > 0)
        {
            try
            {
                int logoCount = logoOrgList.Count;
                int logoHeight;

                // Adjust logo size based on count
                if (logoCount <= 2)
                    logoHeight = 70;
                else if (logoCount <= 5)
                    logoHeight = 50;
                else
                    logoHeight = 40;

                var logosHtml = new StringBuilder();
                logosHtml.Append("<div style='width:100%; margin-top:40px; text-align:center;'>");

                // SME logo first
                logosHtml.Append($"<img src='data:image/jpeg;base64,{logoBase64}' height='{logoHeight}' style='margin-right:10px;' />");

                int count = 1; // SME logo is already added
                foreach (var logo in logoOrgList)
                {
                    string? logox = logo.Organization_Logo;
                    if (!string.IsNullOrEmpty(logox) && logox.Contains("<content>"))
                    {
                        var match = System.Text.RegularExpressions.Regex.Match(logox, @"<content>(.*?)</content>", System.Text.RegularExpressions.RegexOptions.Singleline);
                        if (match.Success)
                        {
                            var base64String = match.Groups[1].Value;
                            logosHtml.Append($"<img src='data:image/png;base64,{base64String}' height='{logoHeight}' style='margin-right:10px;' />");
                            count++;
                            if (count % 5 == 0)
                            {
                                logosHtml.Append("<br/>");
                            }
                        }
                    }
                }
                logosHtml.Append("</div>");
                contractLogoHtml = logosHtml.ToString();
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
        string strAttorneyLetterDate = CommonDAO.ToThaiDateStringCovert(result.Effective_Date ?? DateTime.Now);
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
        if (result.AttorneyFlag == true)
        {
            strAttorney = "ผู้มีอำนาจกระทำการแทนปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ ฉบับลง" + strAttorneyLetterDate + "";

        }
        else
        {
            strAttorney = "";
        }
        #endregion

        //  Font
        //   var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabun.ttf").Replace("\\", "/");
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabun.ttf");
        string fontBase64 = "";
        if (File.Exists(fontPath))
        {
            var bytes = File.ReadAllBytes(fontPath);
            fontBase64 = Convert.ToBase64String(bytes);
        }
        //   Purpose list
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
            font-family: 'TH Sarabun PSK';
              src: url('data:font/truetype;charset=utf-8;base64,{fontBase64}') format('truetype');
            font-weight: normal;
            font-style: normal;
        }}
{contractCss}
        </style>
    </head>
    <body>
     {contractLogoHtml}
    </br>
    </br>
        <div class='t-16 text-center'><B>บันทึกข้อตกลงความร่วมมือ</B></div>
       <div class='t-16 text-center'><B>{CommonDAO.ConvertStringArabicToThaiNumerals(result.ProjectTitle)}</B></div>
        <div class='t-12 text-center'><B>ระหว่าง</B></div>
        <div class='t-12 text-center'><B>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B></div>
        <div class='t-12 text-center'><B>กับ</B></div>
        <div class='t-12 text-center'><B>{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? ""}</B></div>
        <br/>
         <P class='t-12 tab2'>บันทึกข้อตกลงความร่วมมือฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลาง และขนาดย่อม เมื่อ {strSign_Date} ระหว่าง</P>
        <P class='t-12 tab2'><B>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B>  โดย {CommonDAO.ConvertStringArabicToThaiNumerals(result.OSMEP_NAME)} ตำแหน่ง {CommonDAO.ConvertStringArabicToThaiNumerals(result.OSMEP_POSITION)} {CommonDAO.ConvertStringArabicToThaiNumerals(strAttorneyOsmep)} สัญญาฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เลขที่ ๑๒๐ หมู่ ๓ ศูนย์ราชการเฉลิมพระเกียรติ ๘๐ พรรษา ๕ ธันวาคม ๒๕๕๐ (อาคารซี) ชั้น ๒, ๑๐, ๑๑ ถนนแจ้งวัฒนะ แขวงทุ่งสองห้อง เขตหลักสี่ กรุงเทพ ๑๐๒๑๐ ซึ่งต่อไป ในสัญญาฉบับนี้จะเรียกว่า“สสว.” ฝ่ายหนึ่ง กับ</P>
        <P class='t-12 tab2'>“{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgCommonName) ?? ""}” {CommonDAO.ConvertStringArabicToThaiNumerals(result.CP_S_NAME)} ตำแหน่ง {CommonDAO.ConvertStringArabicToThaiNumerals(result.CP_S_POSITION)} {CommonDAO.ConvertStringArabicToThaiNumerals(strAttorney)} สำนักงานตั้งอยู่เลขที่ {CommonDAO.ConvertStringArabicToThaiNumerals(result.Office_Loc)} ซึ่งต่อไปในสัญญาฉบับนี้จะเรียกว่า “{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? ""}” อีกฝ่ายหนึ่ง</P>
        <P class='t-12 tab0'><b>วัตถุประสงค์ของความร่วมมือ</b></P>
        <P class='t-12 tab2'>ทั้งสองฝ่ายมีความประสงค์ที่จะร่วมมือกันเพื่อดำเนินการภายใต้โครงการ {CommonDAO.ConvertStringArabicToThaiNumerals(result.ProjectTitle)} ซึ่งในบันทึกข้อตกลงฉบับนี้ต่อไปจะเรียกว่า “โครงการ” โดยมีรายละเอียดโครงการแผนการดำเนินงาน แผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) และบรรดาเอกสารแนบท้ายบันทึกข้อตกลงฉบับนี้ ซึ่งให้ถือเป็นส่วนหนึ่งของบันทึกข้อตกลงฉบับนี้ มีระยะเวลา ตั้งแต่วันที่ {strStart_Date} จนถึงวันที่ {strEnd_Date} โดยมีวัตถุประสงค์ ในการดำเนินโครงการ ดังนี้</P>
    {(purposeList != null && purposeList.Count > 0
    ? string.Join("", purposeList.Select((p, i) =>
        $"<div class='t-12 tab2'>{CommonDAO.ConvertStringArabicToThaiNumerals(p.Detail)}</div>"))
    : "")}  
    <P class='t-12 tab2'><b>ข้อ ๑ ขอบเขตความร่วมมือของ “สสว.”</b></P>
        <P class='t-12 tab3'>๑.๑ ตกลงร่วมดำเนินการโครงการโดยสนับสนุนงบประมาณ จำนวน {result.Contract_Value?.ToString("N2") ?? "0.00"} บาท  ( {strContract_Value} ) ซึ่งได้รวมภาษีมูลค่าเพิ่ม ตลอดจนค่าภาษีอากรอื่น ๆ แล้วให้กับ “{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? ""}” และการใช้จ่ายเงินให้เป็นไปตามแผน การจ่ายเงินตามเอกสารแนบท้ายบันทึกข้อตกลงฉบับนี้</P>
        <P class='t-12 tab3'>๑.๒ ประสานการดำเนินโครงการ เพื่อให้บรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์</P>
        <P class='t-12 tab3'>๑.๓ กำกับ ติดตามและประเมินผลการดำเนินงานของโครงการ</P>
        <P class='t-12 tab2'><b>ข้อ ๒ ขอบเขตความร่วมมือของ “{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? ""}”</b></P>
        <P class='t-12 tab3'>๒.๑ ตกลงที่จะร่วมดำเนินการโครงการตามวัตถุประสงค์ของการโครงการ 
    และขอบเขต การดำเนินการตามรายละเอียดโครงการ แผนการดำเนินการ และแผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) ที่แนบท้ายบันทึกข้อตกลงฉบับนี้</P>
        <P class='t-12 tab3'>๒.๒ ต้องดำเนินโครงการ ปฏิบัติตามแผนการดำเนินงาน แผนการใช้จ่ายเงิน (หรืออาจมีคู่มือ การดำเนินโครงการก็ได้) อย่างเคร่งครัดและให้แล้วเสร็จภายในระยะเวลาโครงการ</P>
        <P class='t-12 tab3'>๒.๓ ต้องประสานการดำเนินโครงการ เพื่อให้โครงการบรรลุวัตถุประสงค์ เป้าหมายผลผลิต และผลลัพธ์</P>
        <P class='t-12 tab3'>๒.๔ ต้องให้ความร่วมมือกับ สสว.ในการกำกับ ติดตามและประเมินผลการ ดำเนินงานของโครงการ</P>
        <P class='t-12 tab2'><b>ข้อ ๓ อื่น ๆ</b></P>
        <P class='t-12 tab3'>๓.๑ หากฝ่ายใดฝ่ายหนึ่งประสงค์จะขอแก้ไข เปลี่ยนแปลง ขยายระยะเวลาของโครงการ จะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษร และต้องได้รับความยินยอมเป็นลาย ลักษณ์อักษรจากอีกฝ่ายหนึ่ง และต้องทำบันทึกข้อตกลงแก้ไข เปลี่ยนแปลง ขยายระยะเวลา เพื่อลงนาม ยินยอมทั้งสองฝ่าย</P>

    <P class='t-12 tab3'>๓.๒ หากฝ่ายใดฝ่ายหนึ่งประสงค์จะขอบอกเลิกบันทึกข้อตกลงความร่วมมือก่อนครบ 
    กำหนดระยะเวลาดำเนินโครงการจะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษรไม่น้อยกว่า ๓๐ วัน และต้องได้รับความยินยอมเป็นลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และ “{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? ""}” จะต้องคืนเงินในส่วน ที่ยังไม่ได้ใช้จ่ายหรือส่วนที่เหลือทั้งหมดพร้อมดอกผล (ถ้ามี) ให้แก่ สสว. ภายใน ๑๕ วัน นับจากวันที่ได้รับ หนังสือของฝ่ายที่ยินยอมให้บอกเลิก</P>

    <P class='t-12 tab3'>๓.๓ สสว. อาจบอกเลิกบันทึกข้อตกลงความร่วมมือได้ทันที หากตรวจสอบ หรือปรากฏ ข้อเท็จจริงว่า การใช้จ่ายเงินของ “{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? ""}” ไม่เป็นไปตามวัตถุประสงค์ของโครงการ แผนการดำเนินงาน และแผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) ทั้งมีสิทธิเรียกเงินคงเหลือพร้อมดอกผล (ถ้ามี) คืนทั้งหมดได้ทันที</P>
        <P class='t-12 tab3'>๓.๔ ทรัพย์สินใด ๆ และ/หรือ สิทธิใด ๆ ที่ได้มาจากเงินสนับสนุนตามบันทึกข้อตกลงฉบับนี้ เมื่อสิ้นสุดโครงการให้ตกได้แก่ สสว. ทั้งสิ้น เว้นแต่ สสว. จะกำหนดให้เป็นอย่างอื่น</P>
        <P class='t-12 tab3'>๓.๕ “{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? "-"}” ต้องไม่ดำเนินการในลักษณะการจ้างเหมา กับหน่วยงาน องค์กร หรือบุคคลอื่น ๆ ยกเว้นกรณีการจัดหา จัดจ้าง เป็นกิจกรรมหรือเป็นเรื่อง ๆ</P>
        <P class='t-12 tab3'>๓.๖ ในกรณีที่การดำเนินการตามบันทึกข้อตกลงฉบับนี้ เกี่ยวข้องกับข้อมูลส่วนบุคคล และ การคุ้มครองทรัพย์สินทางปัญญา “{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? "-"}” จะต้องปฏิบัติตามกฎหมายว่าด้วยการคุ้มครองข้อมูล ส่วนบุคคลและการคุ้มครองทรัพย์สินทางปัญญาอย่างเคร่งครัด และหากเกิดความเสียหายหรือมีการฟ้อง ร้องใดๆ “{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? "-"}” จะต้องเป็นผู้รับผิดชอบต่อการละเมิดบทบัญญัติแห่งกฎหมายดังกล่าว แต่เพียงฝ่ายเดียวโดยสิ้นเชิง</P>
        <P class='t-12 tab3'>บันทึกข้อตกลงนี้ทำขึ้นเป็นบันทึกข้อตกลงอิเล็กทรอนิกส์ คู่ตกลงได้อ่าน เข้าใจเงื่อนไข และยอมรับเงื่อนไข และได้ยืนยันว่าเป็นผู้มีอำนาจลงนามในบันทึกข้อตกลงจึงได้ลงลายมืออิเล็กทรอนิกส์พร้อมทั้งประทับตรา (ถ้ามี) ในบันทึกข้อตกลงไว้ และต่างฝ่ายต่างยึดถือไว้ฝ่ายละหนึ่งฉบับในระบบของตน</P>


    </br>
    </br>
    {signatoryTableHtml}
       <P class='t-12 tab3'>ข้าพเจ้าขอรับรองว่า ทั้งสองฝ่ายได้ลงนามในบันทึกข้อตกลงโดยวิธีการอิเล็กทรอนิกส์ เพื่อแสดงเจตนาของคู่ตกลงแล้ว ข้าพเจ้าจึงได้ลงลายมือชื่ออิเล็กทรอนิกส์รับรองเป็นพยานในบันทึกข้อตกลงพร้อมนี้</P>

    {signatoryTableHtmlWitnesses}
    </body>
    </html>
    ";



        return html;
    }


    #endregion old
}
