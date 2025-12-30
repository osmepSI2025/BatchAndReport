using BatchAndReport.DAO;
using BatchAndReport.Models;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using System.Text;
using System.Threading.Tasks;

public class WordEContract_MemorandumInWritingService
{
    private readonly WordServiceSetting _w;
    private readonly E_ContractReportDAO _eContractReportDAO;
    private readonly IConverter _pdfConverter; 
    public WordEContract_MemorandumInWritingService(WordServiceSetting ws
            , E_ContractReportDAO eContractReportDAO
        , IConverter pdfConverter
        )
    {
        _w = ws;
        _eContractReportDAO = eContractReportDAO;
        _pdfConverter = pdfConverter;
    }
    #region 4.1.1.2.3.บันทึกข้อตกลงความเข้าใจ
    public async Task<byte[]> OnGetWordContact_MemorandumInWritingService(string id)
    {
        var result = await _eContractReportDAO.GetMOUAsync(id);

        if (result == null)
        {
            throw new Exception("ไม่พบข้อมูลบันทึกข้อตกลงความเข้าใจ");
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

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม โดย "+result.OrgCommonName+ " เลขที่ ๑๒๐ หมู่ ๓ ศูนย์ราชการเฉลิมพระเกียรติ ๘๐ พรรษา ๕ ธันวาคม ๒๕๕๐ (อาคารซี) ชั้น ๒, ๑๐, ๑๑ ถนนแจ้งวัฒนะ แขวงทุ่งสองห้อง เขตหลักสี่ กรุงเทพ ๑๐๒๑๐ ซึ่งต่อไป ในสัญญาฉบับนี้จะเรียกว่า“สสว.”ฝ่ายหนึ่ง กับ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("“ชื่อเต็มของหน่วยงาน” โดย "+result.Requestor+" ตำแหน่ง "+result.RequestorPosition+ "ผู้มีอำนาจกระทำการแทนปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ ฉบับลง"+ strDateTH + "สำนักงานตั้งอยู่เลขที่ ซึ่งต่อไปในสัญญาฉบับนี้จะเรียกว่า “  ” อีกฝ่ายหนึ่ง", null, "32"));
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

    public async Task<string> OnGetWordContact_MemorandumInWritingService_HtmlToPDF(string id, string typeContact)
    {
        var result = await _eContractReportDAO.GetMOAAsync(id);

        if (result == null)
        {
            throw new Exception("ไม่พบข้อมูลบันทึกข้อตกลงความเข้าใจ");
        }

        // Logo
        string strContract_Value = CommonDAO.NumberToThaiText(result.Contract_Value ?? 0);
        string strSign_Date = CommonDAO.ToThaiDateStringCovert(result.Sign_Date ?? DateTime.Now);
        string strStart_Date = CommonDAO.ToThaiDateStringCovert(result.Start_Date ?? DateTime.Now);
        string strEnd_Date = CommonDAO.ToThaiDateStringCovert(result.End_Date ?? DateTime.Now);

        var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
        string logoBase64 = "";
        if (System.IO.File.Exists(logoPath))
        {
            var bytes = System.IO.File.ReadAllBytes(logoPath);
            logoBase64 = Convert.ToBase64String(bytes);
        }
        string contractLogoHtml = "";
        var logoOrgList = await _eContractReportDAO.Getsp_GetOrganizationLogosAsync(id, "MOA");
        // Read CSS file content
        var cssPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "css", "contract.css");
        string contractCss = "";
        if (File.Exists(cssPath))
        {
            contractCss = File.ReadAllText(cssPath, Encoding.UTF8);
        }
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
            strAttorney = "ผู้มีอำนาจกระทำการแทน ปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ ฉบับลง" + strAttorneyLetterDate + "";

        }
        else
        {
            strAttorney = "";
        }
        #endregion

        // Font
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabun.ttf");
        string fontBase64 = "";
        if (File.Exists(fontPath))
        {
            var bytes = File.ReadAllBytes(fontPath);
            fontBase64 = Convert.ToBase64String(bytes);
        }

        // Purpose list
        var purposeList = await _eContractReportDAO.GetMOAPoposeAsync(id);


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

        #region ScopeOfMemorandum
        var xScopeOfMemorandum = await _eContractReportDAO.GetScopeOfMemorandumAsync(id, "MOA");

        var xScopeOfMemorandum_OSMEP = xScopeOfMemorandum.Where(e => e.Owner == "OSMEP").ToList();
        var xScopeOfMemorandum_CP = xScopeOfMemorandum.Where(e => e.Owner == "CP").ToList();
        #endregion

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
    <div class='t-14 text-center'><B>บันทึกข้อตกลงความเข้าใจ</B></div>
   <div class='t-14 text-center'><B> {CommonDAO.ConvertStringArabicToThaiNumerals(result.ProjectTitle)}</B></div>
    <div class='t-12 text-center'><B>ระหว่าง</B></div>
    <div class='t-12 text-center'><B>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B></div>
    <div class='t-12 text-center'><B>กับ</B></div>
    <div class='t-12 text-center'><B>{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? ""}</B></div>
    <br/>
     <P class='t-12 tab2'>บันทึกข้อตกลงความเข้าใจฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจ ขนาดกลางและขนาดย่อม เมื่อ {strSign_Date} ระหว่าง</P>
    <P class='t-12 tab2'><B>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B>  โดย {result.OSMEP_NAME} ตำแหน่ง {CommonDAO.ConvertStringArabicToThaiNumerals(result.OSMEP_POSITION)} {CommonDAO.ConvertStringArabicToThaiNumerals(strAttorneyOsmep)} สำนักงานตั้งอยู่เลขที่ ๑๒๐ หมู่ ๓ ศูนย์ราชการเฉลิมพระเกียรติ ๘๐ พรรษา ๕ ธันวาคม ๒๕๕๐ (อาคารซี) ชั้น ๒, ๑๐, ๑๑ ถนนแจ้งวัฒนะ แขวงทุ่งสองห้อง เขตหลักสี่ กรุงเทพ ๑๐๒๑๐ ซึ่งต่อไป ในสัญญาฉบับนี้จะเรียกว่า“สสว.” ฝ่ายหนึ่ง กับ</P>
    <P class='t-12 tab2'><B>“{result.OrgCommonName ?? ""}”</B> โดย {result.CP_S_NAME} ตำแหน่ง {result.CP_S_POSITION} {CommonDAO.ConvertStringArabicToThaiNumerals(strAttorney)} สำนักงานตั้งอยู่เลขที่ {CommonDAO.ConvertStringArabicToThaiNumerals(result.Office_Loc)} ซึ่งต่อไปในสัญญาฉบับนี้จะเรียกว่า “{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? ""}” อีกฝ่ายหนึ่ง</P>
    <P class='t-12 tab1'><B>วัตถุประสงค์ของความเข้าใจ</B></P>
    <P class='t-12 tab2'>ทั้งสองฝ่ายมีความประสงค์ที่จัดทำบันทึกความเข้าใจ {result.ProjectTitle}  โดยมีรายละเอียดและบรรดาเอกสารแนบท้ายบันทึกข้อตกลงฉบับนี้ ซึ่งให้ถือเป็นส่วนหนึ่งของบันทึกข้อตกลงฉบับนี้ มีระยะเวลา ตั้งแต่วันที่ {strStart_Date} จนถึงวันที่ {strEnd_Date} โดยมีวัตถุประสงค์ ในการดำเนินโครงการ ดังนี้</P>
{(purposeList != null && purposeList.Count > 0
    ? string.Join("", purposeList.Select((p, i) =>
        $"<div class='t-12 tab2'>{CommonDAO.ConvertStringArabicToThaiNumerals(p.Detail)}</div>"))
    : "")}  

  <P class='t-12 tab2'><b>ข้อ ๑ ขอบเขตความเข้าใจของ “สสว.”</b></P>
    {(xScopeOfMemorandum_OSMEP != null && xScopeOfMemorandum_OSMEP.Count > 0
    ? string.Join("", xScopeOfMemorandum_OSMEP.Select((p, i) =>
        $"<div class='t-12 tab2'>{CommonDAO.ConvertStringArabicToThaiNumerals(p.Detail)}</div>"))
    : "")}  
    <P class='t-12 tab2'><b>ข้อ ๒ ขอบเขตความเข้าใจของ “{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? ""}”</b></P>
        {(xScopeOfMemorandum_CP != null && xScopeOfMemorandum_CP.Count > 0
    ? string.Join("", xScopeOfMemorandum_CP.Select((p, i) =>
        $"<div class='t-12 tab2'>{CommonDAO.ConvertStringArabicToThaiNumerals(p.Detail)}</div>"))
    : "")}  

    <P class='t-12 tab2'><b>ข้อ ๓ อื่น ๆ</b></P>
    <P class='t-12 tab3'>๓.๑ หากฝ่ายใดฝ่ายหนึ่งประสงค์จะขอแก้ไข เปลี่ยนแปลง ขยายระยะเวลา ของโครงการ จะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษร และต้องได้ รับความยินยอมเป็นลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และต้องทำ บันทึกข้อตกลงแก้ไข เปลี่ยนแปลง ขยายระยะเวลา เพื่อลงนามยินยอมทั้งสองฝ่าย</P>
    
<P class='t-12 tab3'>๓.๒ หากฝ่ายใดฝ่ายหนึ่งประสงค์จะขอบอกเลิกบันทึกข้อตกลงความเข้าใจ ก่อนครบกำหนด ระยะเวลาดำเนินโครงการจะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่ง ได้ทราบเป็นลายลักษณ์อักษรไม่น้อยกว่า ๓๐ วัน และต้องได้รับความยินยอมเป็นลายลักษณ์ อักษรจากอีกฝ่ายหนึ่ง และ “{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? ""}” จะต้องคืนเงินในส่วน ที่ยังไม่ได้ใช้จ่ายหรือส่วนที่เหลือทั้งหมดพร้อมดอกผล (ถ้ามี) ให้แก่ สสว. ภายใน ๑๕ วัน นับจากวันที่ได้รับหนังสือของฝ่ายที่ยินยอมให้บอกเลิก</P>
 
<P class='t-12 tab3'>๓.๓ สสว. อาจบอกเลิกบันทึกข้อตกลงความเข้าใจได้ทันที หากตรวจสอบ หรือปรากฏข้อเท็จจริงว่า การใช้จ่ายเงินของ “{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? ""}” ไม่เป็นไปตามวัตถุประสงค์ ของโครงการ แผนการดำเนินงาน และแผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) ทั้งมีสิทธิเรียกเงินคงเหลือพร้อมดอกผล (ถ้ามี) คืนทั้งหมดได้ทันที</P>
    <P class='t-12 tab3'>๓.๔ ทรัพย์สินใด ๆ และ/หรือ สิทธิใด ๆ ที่ได้มาจากเงินสนับสนุนตาม บันทึกข้อตกลงฉบับนี้ เมื่อสิ้นสุดโครงการให้ตกได้แก่ สสว. ทั้งสิ้น เว้นแต่ สสว. จะกำหนดให้เป็นอย่างอื่น</P>
    <P class='t-12 tab3'>๓.๕ “{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? ""}” ต้องไม่ดำเนินการในลักษณะการจ้างเหมา กับหน่วยงาน องค์กร หรือบุคคลอื่น ๆ ยกเว้นกรณีการจัดหา จัดจ้าง เป็นกิจกรรมหรือเป็นเรื่อง ๆ</P>
    <P class='t-12 tab3'>๓.๖ ในกรณีที่การดำเนินการตามบันทึกข้อตกลงฉบับนี้ เกี่ยวข้องกับ ข้อมูลส่วนบุคคล และการคุ้มครองทรัพย์สินทางปัญญา “{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? ""}” จะต้องปฏิบัติ ตามกฎหมาย ว่าด้วยการคุ้มครอง ข้อมูลส่วนบุคคลและ การคุ้มครองทรัพย์สินทางปัญญา อย่างเคร่งครัด และหากเกิดความเสียหายหรือมีการฟ้องร้องใดๆ “{CommonDAO.ConvertStringArabicToThaiNumerals(result.OrgName) ?? ""}” จะต้องเป็นผู้รับผิดชอบ ต่อการละเมิดบทบัญญัติแห่งกฎหมายดังกล่าว แต่เพียงฝ่ายเดียว โดยสิ้นเชิง</P>
    <P class='t-12 tab3'>บันทึกความเข้าใจนี้ทำขึ้นเป็นบันทึกความเข้าใจทางอิเล็กทรอนิกส์ คู่ตกลงได้อ่าน เข้าใจเงื่อนไข และยอมรับเงื่อนไข และได้ยืนยันว่าเป็นผู้มีอำนาจลงนามในบันทึกความเข้าใจ จึงได้ลงลายมืออิเล็กทรอนิกส์พร้อมทั้งประทับตรา (ถ้ามี) ในบันทึกความเข้าใจไว้ และต่างฝ่ายต่างยึดถือไว้ฝ่ายละหนึ่งฉบับในระบบของตน  </P>


</br>
</br>
{signatoryTableHtml}

    <P class='t-12 tab3'>ข้าพเจ้าขอรับรองว่า ทั้งสองฝ่ายได้ลงนามในบันทึกความเข้าใจโดยวิธีการอิเล็กทรอนิกส์ เพื่อแสดงเจตนาของคู่ตกลงแล้ว ข้าพเจ้าจึงได้ลงลายมือชื่ออิเล็กทรอนิกส์รับรองเป็นพยานในบันทึกความเข้าใจพร้อมนี้</P>

{signatoryTableHtmlWitnesses}
</body>
</html>
";

      
        return html;
    }
    #endregion
}
