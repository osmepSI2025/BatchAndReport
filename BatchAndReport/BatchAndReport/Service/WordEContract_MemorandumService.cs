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
    public async Task<byte[]> OnGetWordContact_MemorandumService(string id)
    {
        var result = await _eContractReportDAO.GetMOUAsync(id);

        if (result == null)
        {
            throw new Exception("ไม่พบข้อมูลบันทึกข้อตกลงความร่วมมือ");
        }
        else 
        {
            var stream = new MemoryStream();

            using (var wordDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();

                // Styles
                var stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylePart.Styles = WordServiceSetting.CreateDefaultStyles();

                var body = mainPart.Document.AppendChild(new Body());
                // 1. Logo (centered)
                var imagePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");

                // Add image part and feed image data
                var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg, "rIdLogo");
                using (var imgStream = File.OpenRead(imagePath))
                {
                    imagePart.FeedData(imgStream);
                }

                // --- 1. Top Row: Logo left, Contract code box right ---
                var topTable = new Table(
                new TableProperties(
                 new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                 new TableBorders(
                 new TopBorder { Val = BorderValues.None },
                 new BottomBorder { Val = BorderValues.None },
                 new LeftBorder { Val = BorderValues.None },
                 new RightBorder { Val = BorderValues.None },
                 new InsideHorizontalBorder { Val = BorderValues.None },
                 new InsideVerticalBorder { Val = BorderValues.None }
                 )
                ),
                new TableRow(
                 // Logo cell
                 new TableCell(
                 new TableCellProperties(
                 new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "60" }
                 ),
                 new Paragraph(
                 new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
                 // Use your logo image here
                 WordServiceSetting.CreateImage(
                 mainPart.GetIdOfPart(imagePart),
                 240, 80
                 )
                 )
                 ),
                 // Contract code box cell
                 new TableCell(
                 new TableCellProperties(
                 new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "40" },
                 new TableCellBorders(

                 )
                 ),
                 new Paragraph(
                 new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                 WordServiceSetting.CreateImage(
                 mainPart.GetIdOfPart(imagePart),
                 240, 80
                 )
                 )
                 )
                )
                );
                body.AppendChild(topTable);

                // --- 2. Titles ---
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("บันทึกข้อตกลงความร่วมมือ", "44"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("โครงการ"+ result.ProjectTitle , "44"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("ระหว่าง", "32"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม", "32"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("กับ", "36"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph(result.OrgCommonName ?? "", "36"));


                // --- 3. Main contract body ---
                var strDateTH = CommonDAO.ToThaiDateStringCovert(result.Sign_Date ?? DateTime.Now);
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(บันทึกข้อตกลงความร่วมมือฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม" +
                        "เมื่อ"+ strDateTH + " ระหว่าง", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม โดย "+result.OrgCommonName+" เลขที่ 120 หมู่ 3 ศูนย์ราชการเฉลิมพระเกียรติ 80 พรรษา 5 ธันวาคม 2550. (อาคารซี) ชั้น 2, 10, 11 ถนนแจ้งวัฒนะ แขวงทุ่งสองห้อง เขตหลักสี่ กรุงเทพ 10210 ซึ่งต่อไป ในสัญญาฉบับนี้จะเรียกว่า“สสว.”ฝ่ายหนึ่ง กับ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("“ชื่อเต็มของหน่วยงาน” โดย "+result.Requestor+" ตำแหน่ง."+result.RequestorPosition+ ".ผู้มีอำนาจกระทำการแทนปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ ฉบับลง"+ strDateTH + "สำนักงานตั้งอยู่เลขที่ ซึ่งต่อไปในสัญญาฉบับนี้จะเรียกว่า “  ” อีกฝ่ายหนึ่ง", null, "32"));
                body.AppendChild(WordServiceSetting.JustifiedParagraph("วัตถุประสงค์ของความร่วมมือ", "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ทั้งสองฝ่ายมีความประสงค์ที่จะร่วมมือกันเพื่อดำเนินการภายใต้โครงการ (ชื่อโครงการที่ระบุไว้ข้างต้น) ซึ่งในบันทึกข้อตกลงฉบับนี้ต่อไปจะเรียกว่า “โครงการ” โดยมีรายละเอียดโครงการแผนการดำเนินงาน แผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) และบรรดาเอกสารแนบท้ายบันทึกข้อตกลงฉบับนี้ ซึ่งให้ถือเป็นส่วนหนึ่งของบันทึกข้อตกลงฉบับนี้ มีระยะเวลา" +
                         "ตั้งแต่วันที่ " + CommonDAO.ToThaiDateStringCovert(result.Start_Date ?? DateTime.Now) +
                    " จนถึงวันที่" +
                    CommonDAO.ToThaiDateStringCovert(result.End_Date ?? DateTime.Now) +
                    "โดยมีวัตถุประสงค์ในการดำเนินโครงการ ดังนี้",
                    null, "32"));

                var purposeList = await _eContractReportDAO.GetMOUPoposeAsync(id);
                if (purposeList.Count != 0 && purposeList != null)
                {
                    int row = 1;
                    foreach (var purpose in purposeList)
                    {
                        body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs(row.ToString() + "• " + purpose.Detail, null, "32"));
                        row++;
                    }
                }



                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 1 ขอบเขตความร่วมมือของ “สสว.”", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("1.1 ตกลงร่วมดำเนินการโครงการโดยสนับสนุนงบประมาณ " +
                                  "จำนวน " + result.Contract_Value + " บาท ( " + CommonDAO.NumberToThaiText(result.Contract_Value ?? 0) + " ) " +
                    " ซึ่งได้รวมภาษีมูลค่าเพิ่ม ตลอดจนค่าภาษีอากรอื่น ๆ แล้วให้กับ “ชื่อหน่วยร่วม”และการใช้จ่ายเงินให้เป็นไปตามแผนการจ่ายเงินตามเอกสารแนบท้ายบันทึกข้อตกลงฉบับนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("1.2 ประสานการดำเนินโครงการ เพื่อให้บรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("1.3 กำกับ ติดตามและประเมินผลการดำเนินงานของโครงการ", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 2 ขอบเขตความร่วมมือของ “ชื่อหน่วยร่วม”", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("2.1 ตกลงที่จะร่วมดำเนินการโครงการตามวัตถุประสงค์ของการโครงการและขอบเขตการดำเนินการตามรายละเอียดโครงการ แผนการดำเนินการ และแผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) ที่แนบท้ายบันทึกข้อตกลงฉบับนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("2.2 ต้องดำเนินโครงการ ปฏิบัติตามแผนการดำเนินงาน แผนการใช้จ่ายเงิน (หรืออาจมีคู่มือการดำเนินโครงการก็ได้) อย่างเคร่งครัดและให้แล้วเสร็จภายในระยะเวลาโครงการ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("2.3 ต้องประสานการดำเนินโครงการ เพื่อให้โครงการบรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("2.4 ต้องให้ความร่วมมือกับ สสว. ในการกำกับ ติดตามและประเมินผลการดำเนินงานของโครงการ", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 3 อื่น ๆ", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("3.1 หากฝ่ายใดฝ่ายหนึ่งประสงค์จะขอแก้ไข เปลี่ยนแปลง ขยายระยะเวลาของโครงการ จะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษร และต้องได้รับความยินยอมเป็น ลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และต้องทำบันทึกข้อตกลงแก้ไข เปลี่ยนแปลง ขยายระยะเวลา เพื่อลงนามยินยอมทั้งสองฝ่าย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("3.2 หากฝ่ายใดฝ่ายหนึ่งประสงค์จะขอบอกเลิกบันทึกข้อตกลงความร่วมมือก่อนครบกำหนดระยะเวลาดำเนินโครงการจะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษรไม่น้อยกว่า 30 วัน และต้องได้รับความยินยอมเป็นลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และ “ชื่อหน่วยร่วม” จะต้องคืนเงินในส่วนที่ยังไม่ได้ใช้จ่ายหรือส่วนที่เหลือทั้งหมดพร้อมดอกผล (ถ้ามี) ให้แก่ สสว. ภายใน 15 วัน นับจากวันที่ได้รับหนังสือของฝ่ายที่ยินยอมให้บอกเลิก", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("3.3 สสว. อาจบอกเลิกบันทึกข้อตกลงความร่วมมือได้ทันที หากตรวจสอบ หรือปรากฏข้อเท็จจริงว่า การใช้จ่ายเงินของ “ชื่อหน่วยร่วม” ไม่เป็นไปตามวัตถุประสงค์ของโครงการ แผนการดำเนินงาน และแผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) ทั้งมีสิทธิเรียกเงินคงเหลือพร้อมดอกผล (ถ้ามี) คืนทั้งหมดได้ทันที", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("3.4 ทรัพย์สินใด ๆ และ/หรือ สิทธิใด ๆ ที่ได้มาจากเงินสนับสนุนตามบันทึกข้อตกลงฉบับนี้ เมื่อสิ้นสุดโครงการให้ตกได้แก่ สสว. ทั้งสิ้น เว้นแต่ สสว. จะกำหนดให้เป็นอย่างอื่น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("3.5 “ชื่อหน่วยร่วม” ต้องไม่ดำเนินการในลักษณะการจ้างเหมา กับหน่วยงาน องค์กร หรือบุคคลอื่น ๆ ยกเว้นกรณีการจัดหา จัดจ้าง เป็นกิจกรรมหรือเป็นเรื่อง ๆ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("3.6 ในกรณีที่การดำเนินการตามบันทึกข้อตกลงฉบับนี้ เกี่ยวข้องกับข้อมูลส่วนบุคคลและการคุ้มครองทรัพย์สินทางปัญญา “ชื่อหน่วยร่วม” จะต้องปฏิบัติตามกฎหมายว่าด้วยการคุ้มครองข้อมูลส่วนบุคคลและการคุ้มครองทรัพย์สินทางปัญญาอย่างเคร่งครัด และหากเกิดความเสียหายหรือมีการฟ้องร้องใด ๆ “ชื่อหน่วยร่วม” จะต้องเป็นผู้รับผิดชอบต่อการละเมิดบทบัญญัติแห่งกฎหมายดังกล่าวแต่เพียงฝ่ายเดียวโดยสิ้นเชิง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("บันทึกข้อตกลงความร่วมมือฉบับนี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน ทั้งสองฝ่ายได้อ่านและเข้าใจข้อความโดยละเอียดแล้ว จึงได้ลงลายมือชื่อพร้อมประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยานและยึดถือไว้ฝ่ายละฉบับ", null, "32"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());


                // --- 6. Signature lines ---
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม
                body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ)...................................................."));
                body.AppendChild(WordServiceSetting.CenteredParagraph("("+result.OSMEP_Signer??""+")"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // หน่วยงานร่วม
                body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ)...................................................."));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(" + result.OSMEP_Witness??"" + ")"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("ชื่อเต็มหน่วยงาน"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // พยาน 1
                body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ)....................................................พยาน"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(" + result.Contract_Signer??"" + ")"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // พยาน 2
                body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ)....................................................พยาน"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(" + result.Contract_Witness??"" + ")"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // --- 7. Add header/footer if needed ---
                WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);
            }
            stream.Position = 0;
            return stream.ToArray();
        }

    }

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
            strAttorneyOsmep = "ผู้มีอำนาจกระทำการแทนปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ เลขคำสั่งที่ " + result.AttorneyLetterNumber + " ฉบับลงวันที่ " + strAttorneyLetterDate + "";

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
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");

        // Purpose list
        var purposeList = await _eContractReportDAO.GetMOUPoposeAsync(id);


        #region signlist 

        var signlist = await _eContractReportDAO.GetSignNameAsync(id, typeContact);
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
        body, p, div, table, th, td {{
            font-family: 'THSarabunNew', Arial, sans-serif !important;
            font-size: 22px;
        }}
        /* แก้การตัดคำไทย: ไม่หั่นกลางคำ, ตัดเมื่อจำเป็น */
        body, p, div {{
            word-break: keep-all;
            overflow-wrap: break-word;
            -webkit-line-break: after-white-space;
            hyphens: none;
        }}
        .t-16 {{
            font-size: 1.5em;
        }}
        .t-18 {{
            font-size: 1.7em;
        }}
        .t-22 {{
            font-size: 1.9em;
        }}
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
        .table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            font-size: 28pt;
        }}
        .table th, .table td {{
            border: 1px solid #000;
            padding: 8px;
        }}
        .sign-double {{ display: flex; }}
        .text-center-right-brake {{
            margin-left: 50%;
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
            font-family: 'THSarabunNew', Arial, sans-serif !important;
        }}
        .logo-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 40px;
        }}
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
    <div class='t-16 text-center'><B>ระหว่าง</B></div>
    <div class='t-22 text-center'><B>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B></div>
    <div class='t-22 text-center'><B>กับ</B></div>
    <div class='t-18 text-center'><B>{result.OrgName ?? ""}</B></div>
    <br/>
     <P class='t-16 tab2'>บันทึกข้อตกลงความร่วมมือฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลาง และขนาดย่อม เมื่อ {strSign_Date} ระหว่าง</P>
    <P class='t-16 tab2'><B>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B>  โดย {result.OSMEP_NAME} ตำแหน่ง {result.OSMEP_POSITION} {strAttorneyOsmep} สำนักงานตั้งอยู่เลขที่ 120 หมู่ 3 ศูนย์ราชการเฉลิมพระเกียรติ 80 พรรษา 5 ธันวาคม 2550. (อาคารซี) ชั้น 2, 10, 11 ถนนแจ้งวัฒนะ แขวงทุ่งสองห้อง เขตหลักสี่ กรุงเทพ 10210 ซึ่งต่อไป ในสัญญาฉบับนี้จะเรียกว่า“สสว.” ฝ่ายหนึ่ง กับ</P>
    <P class='t-16 tab2'>“{result.OrgCommonName ?? ""}” {result.CP_S_NAME} ตำแหน่ง {result.CP_S_POSITION} {strAttorney} สำนักงานตั้งอยู่เลขที่ {result.Office_Loc} ซึ่งต่อไปในสัญญาฉบับนี้จะเรียกว่า “{result.OrgName ?? ""}” อีกฝ่ายหนึ่ง</P>
    <P class='t-16 tab2'>วัตถุประสงค์ของความร่วมมือ</P>
    <P class='t-16 tab2'>ทั้งสองฝ่ายมีความประสงค์ที่จะร่วมมือกันเพื่อดำเนินการภายใต้โครงการ {result.ProjectTitle} ซึ่งในบันทึกข้อตกลงฉบับนี้ต่อไปจะเรียกว่า “โครงการ” โดยมีรายละเอียดโครงการแผนการดำเนินงาน แผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) และบรรดาเอกสารแนบท้ายบันทึกข้อตกลงฉบับนี้ ซึ่งให้ถือเป็นส่วนหนึ่งของบันทึกข้อตกลงฉบับนี้ มีระยะเวลา ตั้งแต่วันที่ {strStart_Date} จนถึงวันที่ {strEnd_Date} โดยมีวัตถุประสงค์ ในการดำเนินโครงการ ดังนี้</P>
{(purposeList != null && purposeList.Count != 0
    ? $"<div class='t-16 tab3'>{string.Join("<br/>", purposeList.Select(p => p.Detail))}</div>"
    : "")}
  <P class='t-16 tab2'><b>ข้อ 1 ขอบเขตความร่วมมือของ “สสว.”</b></P>
    <P class='t-16 tab3'>1.1 ตกลงร่วมดำเนินการโครงการโดยสนับสนุนงบประมาณ จำนวน {result.Contract_Value?.ToString("N2") ?? "0.00"} บาท  ( {strContract_Value} ) ซึ่งได้รวมภาษีมูลค่าเพิ่ม ตลอดจนค่าภาษีอากรอื่น ๆ แล้วให้กับ “{result.OrgName ?? ""}” และการใช้จ่ายเงินให้เป็นไปตามแผน การจ่ายเงินตามเอกสารแนบท้ายบันทึกข้อตกลงฉบับนี้</P>
    <P class='t-16 tab3'>1.2 ประสานการดำเนินโครงการ เพื่อให้บรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์</P>
    <P class='t-16 tab3'>1.3 กำกับ ติดตามและประเมินผลการดำเนินงานของโครงการ</P>
    <P class='t-16 tab2'><b>ข้อ 2 ขอบเขตความร่วมมือของ “{result.OrgName ?? ""}”</b></P>
    <P class='t-16 tab3'>2.1 ตกลงที่จะร่วมดำเนินการโครงการตามวัตถุประสงค์ของการโครงการ 
และขอบเขต การดำเนินการตามรายละเอียดโครงการ แผนการดำเนินการ และแผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) ที่แนบท้ายบันทึกข้อตกลงฉบับนี้</P>
    <P class='t-16 tab3'>2.2 ต้องดำเนินโครงการ ปฏิบัติตามแผนการดำเนินงาน แผนการใช้จ่ายเงิน (หรืออาจมีคู่มือ การดำเนินโครงการก็ได้) อย่างเคร่งครัดและให้แล้วเสร็จภายในระยะเวลาโครงการ</P>
    <P class='t-16 tab3'>2.3 ต้องประสานการดำเนินโครงการ เพื่อให้โครงการบรรลุวัตถุประสงค์ เป้าหมายผลผลิต และผลลัพธ์</P>
    <P class='t-16 tab3'>2.4 ต้องให้ความร่วมมือกับ สสว.ในการกำกับ ติดตามและประเมินผลการ ดำเนินงานของโครงการ</P>
    <P class='t-16 tab2'><b>ข้อ 3 อื่น ๆ</b></P>
    <P class='t-16 tab3'>3.1 หากฝ่ายใดฝ่ายหนึ่งประสงค์จะขอแก้ไข เปลี่ยนแปลง ขยายระยะเวลาของโครงการ จะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษร และต้องได้รับความยินยอมเป็นลาย ลักษณ์อักษรจากอีกฝ่ายหนึ่ง และต้องทำบันทึกข้อตกลงแก้ไข เปลี่ยนแปลง ขยายระยะเวลา เพื่อลงนาม ยินยอมทั้งสองฝ่าย</P>
   
<P class='t-16 tab3'>3.2 หากฝ่ายใดฝ่ายหนึ่งประสงค์จะขอบอกเลิกบันทึกข้อตกลงความร่วมมือก่อนครบ 
กำหนดระยะเวลาดำเนินโครงการจะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษรไม่น้อยกว่า 30 วัน และต้องได้รับความยินยอมเป็นลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และ “{result.OrgName ?? ""}” จะต้องคืนเงินในส่วน ที่ยังไม่ได้ใช้จ่ายหรือส่วนที่เหลือทั้งหมดพร้อมดอกผล (ถ้ามี) ให้แก่ สสว. ภายใน 15 วัน นับจากวันที่ได้รับ หนังสือของฝ่ายที่ยินยอมให้บอกเลิก</P>
 
<P class='t-16 tab3'>3.3 สสว. อาจบอกเลิกบันทึกข้อตกลงความร่วมมือได้ทันที หากตรวจสอบ หรือปรากฏ ข้อเท็จจริงว่า การใช้จ่ายเงินของ “{result.OrgName ?? ""}” ไม่เป็นไปตามวัตถุประสงค์ของโครงการ แผนการดำเนินงาน และแผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) ทั้งมีสิทธิเรียกเงินคงเหลือพร้อมดอกผล (ถ้ามี) คืนทั้งหมดได้ทันที</P>
    <P class='t-16 tab3'>3.4 ทรัพย์สินใด ๆ และ/หรือ สิทธิใด ๆ ที่ได้มาจากเงินสนับสนุนตามบันทึกข้อตกลงฉบับนี้ เมื่อสิ้นสุดโครงการให้ตกได้แก่ สสว. ทั้งสิ้น เว้นแต่ สสว. จะกำหนดให้เป็นอย่างอื่น</P>
    <P class='t-16 tab3'>3.5 “ชื่อหน่วยร่วม” ต้องไม่ดำเนินการในลักษณะการจ้างเหมา กับหน่วยงาน องค์กร หรือบุคคลอื่น ๆ ยกเว้นกรณีการจัดหา จัดจ้าง เป็นกิจกรรมหรือเป็นเรื่อง ๆ</P>
    <P class='t-16 tab3'>3.6 ในกรณีที่การดำเนินการตามบันทึกข้อตกลงฉบับนี้ เกี่ยวข้องกับข้อมูลส่วนบุคคล และ การคุ้มครองทรัพย์สินทางปัญญา “ชื่อหน่วยร่วม” จะต้องปฏิบัติตามกฎหมายว่าด้วยการคุ้มครองข้อมูล ส่วนบุคคลและการคุ้มครองทรัพย์สินทางปัญญาอย่างเคร่งครัด และหากเกิดความเสียหายหรือมีการฟ้อง ร้องใดๆ “ชื่อหน่วยร่วม” จะต้องเป็นผู้รับผิดชอบต่อการละเมิดบทบัญญัติแห่งกฎหมายดังกล่าว แต่เพียงฝ่ายเดียวโดยสิ้นเชิง</P>
    <P class='t-16 tab3'>บันทึกข้อตกลงความร่วมมือฉบับนี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน ทั้งสองฝ่าย ได้อ่านและเข้าใจข้อความโดยละเอียดแล้ว จึงได้ลงลายมือชื่อพร้อมประทับตรา(ถ้ามี) ไว้เป็นสำคัญต่อหน้า พยาน และยึดถือไว้ฝ่ายละฉบับ</P>


</br>
</br>
{signatoryTableHtml}
</body>
</html>
";



        return html;
    }
    #endregion
}
