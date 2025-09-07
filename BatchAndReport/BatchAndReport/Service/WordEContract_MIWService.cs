using BatchAndReport.DAO;
using BatchAndReport.Services;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Commons.Bouncycastle.Crypto;
using iText.Signatures;
using Spire.Doc.Documents;
using System.Text;
using System.Threading.Tasks;
using static SkiaSharp.HarfBuzz.SKShaper;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
public class WordEContract_MIWService
{
    private readonly WordServiceSetting _w;
    private readonly E_ContractReportDAO _eContractReportDAO;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter

    public WordEContract_MIWService(
        WordServiceSetting ws,
        E_ContractReportDAO eContractReportDAO
      , IConverter pdfConverter
    )
    {
        _w = ws;
        _eContractReportDAO = eContractReportDAO;
        _pdfConverter = pdfConverter;
    }
    #region 4.1.1.2.1.สัญญาร่วมดำเนินการ
    public async Task<byte[]> OnGetWordContact_MIWService(string conId)
    {
        var dataResult = await _eContractReportDAO.GetJOAAsync(conId);
        if (dataResult == null)
        {
            throw new Exception("JOA data not found.");
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
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("สัญญาร่วมดำเนินการ", "44"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("โครงการ " + dataResult.Project_Name, "44"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("ระหว่าง", "32"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม", "32"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("กับ", "36"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph(dataResult.Organization ?? "", "36"));


                // --- 3. Main contract body ---
                var strDateTH = CommonDAO.ToThaiDateString(dataResult.Contract_SignDate ?? DateTime.Now);
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สัญญาร่วมดำเนินการฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม " +
                    "เมื่อวันที่ " + strDateTH[0] + " เดือน " + strDateTH[1] + "พ.ศ." + strDateTH[2] + " ระหว่าง", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม โดย " + dataResult.Organization ?? "" + "เลขที่ 120 หมู่ 3 ศูนย์ราชการเฉลิมพระเกียรติ 80 พรรษา 5 ธันวาคม 2550. (อาคารซี) ชั้น 2, 10, 11 ถนนแจ้งวัฒนะ แขวงทุ่งสองห้อง เขตหลักสี่ กรุงเทพ 10210 ซึ่งต่อไป ในสัญญาฉบับนี้จะเรียกว่า“สสว.” ฝ่ายหนึ่ง กับ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs(" " + dataResult.OfficeLoc + " โดย " + dataResult.IssueOwner + " ตำแหน่ง " + dataResult.IssueOwnerPosition + " ผู้มีอำนาจกระทำการแทน ปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ " +
                    "ฉบับลงวันที่ " + strDateTH[0] + " เดือน" + strDateTH[1] + "พ.ศ." + strDateTH[2] + " ซึ่งต่อไปในสัญญาฉบับนี้จะเรียกว่า “ชื่อหน่วยร่วม” อีกฝ่ายหนึ่ง", null, "32"));
                body.AppendChild(WordServiceSetting.JustifiedParagraph("วัตถุประสงค์ตามสัญญาร่วมดำเนินการ", "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs(
                    "คู่สัญญาทั้งสองฝ่ายมีความประสงค์ที่จะร่วมมือกันเพื่อดำเนินการภายใต้โครงการ " + dataResult.Project_Name +
                    " ซึ่งต่อไปในสัญญานี้จะเรียกว่า “โครงการ” โดยมีรายละเอียดโครงการ แผนการดำเนินงาน แผนการใช้จ่ายเงิน (และอื่นๆ เช่น คู่มือดำเนินโครงการ) และบรรดาเอกสารแนบท้ายสัญญาฉบับนี้ ซึ่งให้ถือเป็นส่วนหนึ่งของสัญญาฉบับนี้ มีระยะเวลา" +
                    "ตั้งแต่วันที่ " + CommonDAO.ToThaiDateStringCovert(dataResult.Contract_Start_Date ?? DateTime.Now) +
                    " จนถึงวันที่" +
                    CommonDAO.ToThaiDateStringCovert(dataResult.Contract_End_Date ?? DateTime.Now) +
                    "โดยมีวัตถุประสงค์ในการดำเนินโครงการ ดังนี้",
                    null, "32"));


                var purposeList = await _eContractReportDAO.GetJOAPoposeAsync(conId);
                if (purposeList.Count != 0 && purposeList != null)
                {
                    int row = 1;
                    foreach (var purpose in purposeList)
                    {
                        body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs(row.ToString() + "• " + purpose.Detail, null, "32"));
                        row++;
                    }
                }

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 1 ขอบเขตหน้าที่ของ “สสว.”", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("1.1 ตกลงร่วมดำเนินการโครงการโดยสนับสนุนงบประมาณ " +
                    "จำนวน " + dataResult.Contract_Value + " บาท ( " + CommonDAO.NumberToThaiText(dataResult.Contract_Value ?? 0) + " ) " +
                    "ซึ่งได้รวมภาษีมูลค่าเพิ่ม ตลอดจนค่าภาษีอากรอื่นๆ แล้วให้กับ “ชื่อหน่วยร่วม” และการใช้จ่ายเงินให้เป็นไปตามแผนการจ่ายเงินตามเอกสารแนบท้ายสัญญา", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("1.2 ประสานการดำเนินโครงการ เพื่อให้บรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("1.3 กำกับ ติดตามและประเมินผลการดำเนินงานของโครงการ", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 2 ขอบเขตหน้าที่ของ “ชื่อหน่วยร่วม” ", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("2.1 ตกลงที่จะร่วมดำเนินการโครงการตามวัตถุประสงค์ของการโครงการและขอบเขตการดำเนินการ ตามรายละเอียดโครงการ แผนการดำเนินการ และแผนการใช้จ่ายเงิน (และอื่นๆ เช่น คู่มือดำเนินโครงการ) ที่แนบท้ายสัญญาฉบับนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("2.2 ต้องดำเนินโครงการ ปฏิบัติตามแผนการดำเนินงาน แผนการใช้จ่ายเงิน (หรืออาจมีคู่มือการดำเนินโครงการก็ได้) อย่างเคร่งครัดและให้แล้วเสร็จภายในระยะเวลาโครงการหากไม่ดำเนินโครงการให้แล้วเสร็จตามที่กำหนดยินยอมชำระค่าปรับให้แก่ สสว. ในอัตราร้อยละ 0.1 ของจำนวนงบประมาณที่ได้รับการสนับสนุนทั้งหมดต่อวัน นับถัดจากวันที่กำหนด แล้วเสร็จ และถ้าหากเห็นว่า “ชื่อหน่วยร่วม” ไม่อาจปฏิบัติตามสัญญาต่อไปได้ “ชื่อหน่วยร่วม” ยินยอมให้ สสว. ใช้สิทธิบอกเลิกสัญญาได้ทันที", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("2.3 ต้องประสานการดำเนินโครงการ เพื่อให้บรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("2.4 ต้องให้ความร่วมมือกับ สสว. ในการกำกับ ติดตามและประเมินผลการดำเนินงานของโครงการ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 3 อื่น ๆ", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.1 หากคู่สัญญาฝ่ายใดฝ่ายหนึ่งประสงค์จะขอแก้ไข เปลี่ยนแปลง ขยายระยะเวลาของโครงการ จะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษร และต้องได้รับความยินยอมเป็นลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และต้องทำเอกสารแก้ไข เปลี่ยนแปลง ขยายระยะเวลา เพื่อลงนามยินยอม ทั้งสองฝ่าย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.2 หากคู่สัญญาฝ่ายใดฝ่ายหนึ่งประสงค์จะขอบอกเลิกสัญญาก่อนครบกำหนดระยะเวลาดำเนินโครงการ จะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษรไม่น้อยกว่า 30 วัน และต้องได้รับความยินยอมเป็นลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และ “ชื่อหน่วยร่วม” จะต้องคืนเงินในส่วนที่ยังไม่ได้ใช้จ่ายหรือส่วนที่เหลือทั้งหมดพร้อมดอกผล (ถ้ามี) ให้แก่ สสว. ภายใน 15 วัน นับจากวันที่ได้รับหนังสือของฝ่ายที่ยินยอมให้บอกเลิก", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.3 สสว. อาจบอกเลิกสัญญาได้ทันที หากตรวจสอบ หรือปรากฏข้อเท็จจริงว่า การใช้จ่ายเงินของ “ชื่อหน่วยร่วม” ไม่เป็นไปตามวัตถุประสงค์ของโครงการ แผนการดำเนินงาน และแผนการใช้จ่ายเงิน (และอื่นๆ เช่น คู่มือดำเนินโครงการ) ทั้งมีสิทธิเรียกเงินคงเหลือคืนทั้งหมดพร้อมดอกผล (ถ้ามี) ได้ทันที", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.4 ทรัพย์สินใดๆ และ/หรือ สิทธิใดๆ ที่ได้มาจากเงินสนับสนุนตามสัญญา ร่วมดำเนินการฉบับนี้ เมื่อสิ้นสุดโครงการให้ตกได้แก่ สสว. ทั้งสิ้น เว้นแต่ สสว. จะกำหนดให้เป็นอย่างอื่น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.5 “ชื่อหน่วยร่วม” ต้องไม่ดำเนินการในลักษณะการจ้างเหมา กับหน่วยงาน องค์กร หรือบุคคลอื่นๆ ยกเว้นกรณีการจัดหา จัดจ้าง เป็นกิจกรรมหรือเป็นเรื่อง ๆ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.6 ในกรณีที่การดำเนินการตามสัญญาฉบับนี้ เกี่ยวข้องกับข้อมูลส่วนบุคคล และการคุ้มครองทรัพย์สินทางปัญญา “ชื่อหน่วยร่วม” จะต้องปฏิบัติตามกฎหมายว่าด้วยการคุ้มครอง ข้อมูลส่วนบุคคลและการคุ้มครองทรัพย์สินทางปัญญาอย่างเคร่งครัดและหากเกิดความเสียหายหรือมีการฟ้องร้องใดๆ “ชื่อหน่วยร่วม” จะต้องเป็นผู้รับผิดชอบต่อการละเมิดบทบัญญัติแห่งกฎหมายดังกล่าวแต่เพียงฝ่ายเดียวโดยสิ้นเชิง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สัญญาฉบับนี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน ทั้งสองฝ่ายได้อ่านและเข้าใจข้อความโดยละเอียดแล้ว จึงได้ลงลายมือชื่อพร้อมประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยานและยึดถือไว้ฝ่ายละฉบับ", null, "32"));


                // --- 6. Signature lines ---
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // First signature block: สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม
                body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ)....................................................", "32"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(                              )", "32"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม", "32"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // Second signature block: หน่วยงานร่วม
                body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ)....................................................", "32"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(                              )", "32"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("ชื่อเต็มหน่วยงาน", "32"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // Third signature block
                body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ)....................................................", "32"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(                              )", "32"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // Fourth signature block
                body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ)....................................................", "32"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(                              )", "32"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // --- 7. Add header/footer if needed ---
                WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);
            }
            stream.Position = 0;
            return stream.ToArray();
        }

    }
    #endregion 4.1.1.2.1.สัญญาร่วมดำเนินการ

    public async Task<string> OnGetWordContact_MIWServiceHtmlToPDF(string conId)
    {
        var dataResult = await _eContractReportDAO.GetJOAAsync(conId);
        if (dataResult == null)
            throw new Exception("JOA data not found.");
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
        #region checkมอบอำนาจ
        string strAttorneyLetterDate = CommonDAO.ToArabicDateStringCovert(dataResult.Grant_Date ?? DateTime.Now);
        string strAttorney = "";
        var HtmlAttorney = new StringBuilder();
        if (dataResult.AttorneyFlag == true)
        {
            strAttorney = "ผู้มีอำนาจ กระทำการแทน ปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ ฉบับลงวันที่ " + strAttorneyLetterDate + "";

        }
        else
        {
            strAttorney = "";
        }
        #endregion

        var strDateTH = CommonDAO.ToThaiDateString(dataResult.Contract_SignDate ?? DateTime.Now);
        var purposeList = await _eContractReportDAO.GetJOAPoposeAsync(conId);

        var signatoryHtml = new StringBuilder();
        var companySealHtml = new StringBuilder();
        bool sealAdded = false; // กันซ้ำ

        foreach (var signer in dataResult.Signatories)
        {
            string signatureHtml;
            string companySeal = ""; // กัน warning

            // ► ลายเซ็นรายบุคคล (เดิม)
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
                    signatureHtml = "<div class='t-16 text-center tab1'>(ลงชื่อ....................)</div>";
                }
            }
            else
            {
                signatureHtml = "<div class='t-16 text-center tab1'>(ลงชื่อ....................)</div>";
            }
            // ► ตราประทับ: ให้พิจารณาเมื่อเจอ CP_S เท่านั้น (ไม่เช็ค null/empty ตรง if ชั้นนอก)
            if (!sealAdded && signer?.Signatory_Type == "CP_S")
            {
                if (!string.IsNullOrEmpty(signer.Company_Seal) && signer.Company_Seal.Contains("<content>"))
                {
                    try
                    {
                        var contentStart = signer.Company_Seal.IndexOf("<content>") + "<content>".Length;
                        var contentEnd = signer.Company_Seal.IndexOf("</content>");
                        var base64 = signer.Company_Seal.Substring(contentStart, contentEnd - contentStart);

                        companySeal = $@"
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
                        companySealHtml.AppendLine("<div class='t-16 text-center tab1'>(ตราประทับ บริษัท)</div>");
                        sealAdded = true;
                    }
                }
                else
                {
                    // ไม่มีไฟล์ตรา/ไม่มี <content> ⇒ ใส่ placeholder ครั้งเดียว
                    companySealHtml.AppendLine("<div class='t-16 text-center tab1'>(ตราประทับ บริษัท)</div>");
                    sealAdded = true;
                }
            }

            signatoryHtml.AppendLine($@"
<div class='sign-single-right'>
    {signatureHtml}
    <div class='t-16 text-center tab1'>({signer?.Signatory_Name})</div>
    <div class='t-16 text-center tab1'>{signer?.BU_UNIT}</div>
</div>");
        }

        // ► Fallback: ถ้าจบลูปแล้วยังไม่มีตราประทับ แต่คุณ “ต้องการให้มีอย่างน้อย placeholder 1 ครั้ง”
        if (!sealAdded)
        {
            companySealHtml.AppendLine("<div class='t-16 text-center tab1'>(ตราประทับ บริษัท)</div>");
            sealAdded = true;
        }

        // ► ประกอบผลลัพธ์
        var signatoryWithLogoHtml = new StringBuilder();
        if (companySealHtml.Length > 0) signatoryWithLogoHtml.Append(companySealHtml);
        signatoryWithLogoHtml.Append(signatoryHtml);



        var html = $@"<html>
<head>
    <meta charset='utf-8'>
  
     <style>
        @font-face {{
            font-family: 'THSarabunNew';
            src: url('file:///{fontPath}') format('truetype');
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
        .t-16 {{
            font-size: 1.5em;
        }}
        .t-18 {{
            font-size: 1.7em;
        }}
        .t-22 {{
            font-size: 1.9em;
        }}
        .tab0 {{ text-indent: 0px;     }}
        .tab1 {{ text-indent: 48px;     }}
        .tab2 {{ text-indent: 96px;    }}
        .tab3 {{ text-indent: 144px;    }}
        .tab4 {{ text-indent: 192px;   }}
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
        }}
     .logo-table {{ width: 100%; border-collapse: collapse; margin-top: 40px; }}
        .logo-table td {{ border: none; }}
        p {{
            margin: 0;
            padding: 0;
        }}
    </style>
</head><body>



</br>
    <div class='t-22 text-center'><b>บันทึกข้อตกลง</b></div>
    <div class='t-22 text-center'><b>จ้างเหมาบริการ…………………………………………..</b></div>
    <div class='t-16 text-center'><b>บันทึกข้อตกลงเลขที่ ................/๒๕๖๗</b></div>
 
</br>
    <P class='t-16 tab3'>
        บันทึกข้อตกลงฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เลขที่ ๑๒๐ หมู่ ๓ ศูนย์ราชการเฉลิมพระเกียรติ ๘๐ พรรษา ๕ ธันวาคม ๒๕๕๐ (อาคารซี) ชั้น ๒, ๑๐, ๑๑ ถนนแจ้งวัฒนะ แขวงทุ่งสองห้อง เขตหลักสี่ กรุงเทพมหานคร ๑๐๒๑๐ 
 เมื่อวันที่…………………………………..ระหว่าง สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม 
โดย ……………………………………………….. ซึ่งต่อไปในบันทึกข้อตกลงนี้เรียกว่า “ผู้ว่าจ้าง” ฝ่ายหนึ่ง 
กับ ……………………………….. ผู้ถือบัตรประจำตัวประชาชนเลขที่..................................  
วันออกบัตร ............................. บัตรหมดอายุ ............................... 
อยู่บ้านเลขที่ ..................... หมู่ที่ ..............ถนน.......................... 
ตำบล...................... อำเภอ.................. จังหวัด......................... ๙๕๐๐๐ 
ปรากฏตามเอกสารแนบท้ายบันทึกข้อตกลงนี้ ซึ่งต่อไปในบันทึกข้อตกลงนี้เรียกว่า “ผู้รับจ้าง” อีกฝ่ายหนึ่ง

    </P>
<p class='t-16 tab3'><b>ข้อ ๑.ข้อตกลงว่าจ้าง</b></p>
  <P class='t-16 tab3'>จ้างและผู้รับจ้างตกลงรับจ้างเหมาบริการ………………………………………………..
ตามข้อกำหนดและเงื่อนไขของบันทึกข้อตกลงนี้รวมทั้งเอกสารแนบท้ายบันทึกข้อตกลง
    </P>
    <P class='t-16 tab3'><b>ข้อ ๒.เอกสารแนบท้ายบันทึกข้อตกลง</b>
</P>
    <P class='t-16 tab4'>เอกสารแนบท้ายบันทึกข้อตกลงดังต่อไปนี้ให้ถือเป็นส่วนหนึ่งของบันทึกข้อตกลงนี้
</P>
    <P class='t-16 tab4'>๒.๑ ผนวก ๑ ขอบเขตของงาน (TOR) จำนวน 	.............  หน้า
</P>
    <P class='t-16 tab4'>๒.๒ ผนวก ๒ สำเนาบัตรประจำตัวประชาชน จำนวน  ............ หน้า
</P>
 <P class='t-16 tab4'>๒.๓ ผนวก ๓ สำเนาทะเบียนบ้าน จำนวน  ............ หน้า
</P>
 <P class='t-16 tab4'>๒.๔ ผนวก ๔ ใบเสนอราคา หนังสือประสบการณ์ทำงาน วุฒิการศึกษา จำนวน  ........... หน้า  
</P>
 <P class='t-16 tab4'>ความใดในเอกสารแนบท้ายบันทึกข้อตกลงที่ขัดแย้งกับข้อความในบันทึกข้อตกลงนี้ให้
ใช้ข้อความในบันทึกข้อตกลงนี้บังคับในกรณีที่เอกสารแนบท้ายบันทึกข้อตกลงขัดแย้งกันเอง
ผู้รับจ้างจะต้องปฏิบัติตามคำวินิจฉัยของผู้ว่าจ้าง
ทั้งนี้ ผู้รับจ้างไม่มีสิทธิเรียกร้องค่าเสียหายหรือค่าใช้จ่ายใดๆ ทั้งสิ้น
</P>
    <P class='t-16 tab3'><B>ข้อ ๓. ค่าจ้างและการจ่ายเงิน 
</B></P>
    <p class='t-16 tab4'>ผู้ว่าจ้างตกลงจ่ายและผู้รับจ้างตกลงรับเงินค่าจ้างจำนวนเงิน ...... บาท(...........)
ซึ่งได้รวมภาษีอากรอื่นๆ และค่าใช้จ่ายทั้งปวงด้วยแล้ว โดยกำหนดการจ่ายเงินเป็นรายเดือน จำนวนเงินเดือนละ ........... บาท (...........) 
เมื่อผู้รับจ้างได้ปฏิบัติงานและนำส่งรายงานผลการปฏิบัติงาน และใบบันทึกลงเวลาการปฏิบัติงานในแต่ละเดือนให้แก่ผู้ว่าจ้าง 
ภายในวันที่ ๕ ของเดือนถัดไป นับจากวันสิ้นสุดของงานในแต่ละงวด ยกเว้นงวดสุดท้ายให้ส่งมอบภายในวันที่ ........... 
ซึ่งมีรายละเอียดของงานปรากฏตามเอกสารแนบท้ายบันทึกข้อตกลง และผู้ว่าจ้างได้ตรวจรับงานจ้างไว้โดยครบถ้วนแล้ว</p>
   <P class='t-16 tab4'>ทั้งนี้ หากเดือนใดมีการปฏิบัติงานไม่เต็มเดือนปฏิทิน ให้คิดค่าจ้างเหมาเป็นรายวัน ในอัตราวันละ ........... บาท (...........) </P>
    <P class='t-16 tab4'>การจ่ายเงินตามเงื่อนไขแห่งสัญญานี้ ผู้ว่าจ้างจะโอนเงินเข้าบัญชีเงินฝากธนาคาร ของผู้รับจ้าง 
ชื่อธนาคาร...........สาขา........... ชื่อบัญชี ........... 
เลขที่บัญชี ........... ทั้งนี้ ผู้รับจ้างตกลงเป็นผู้รับภาระเงินค่าธรรมเนียมหรือค่าบริการอื่นใดเกี่ยวกับการโอน
รวมทั้งค่าใช้จ่ายอื่นใด (ถ้ามี) ที่ธนาคารเรียกเก็บ และยินยอมให้มีการหักเงินดังกล่าวจากจำนวนเงินโอนในงวด</P>

  <P class='t-16 tab3'><b>ข้อ ๔.กำหนดเวลาแล้วเสร็จและสิทธิของผู้ว่าจ้างในการบอกเลิกบันทึกข้อตกลง</b></P>
    <p class='t-16 tab4'>ผู้รับจ้างต้องเริ่มทำงานที่รับจ้างภายในวันที่ ........... และจะต้องทำงานให้แล้วเสร็จบริบูรณ์ภายในวันที่ ...........
ตามรายละเอียดเอกสารแนบท้ายบันทึกข้อตกลงนี้ และต้องผ่านการตรวจรับผลการปฏิบัติงานจากผู้ว่าจ้างในแต่ละเดือน
ถ้าผู้รับจ้างมิได้ลงมือทำงานภายในกำหนดเวลา หรือไม่สามารถทำงานให้ครบถ้วนตามเงื่อนไขของบันทึกข้อตกลงนี้
หรือมีเหตุให้เชื่อได้ว่า ผู้รับจ้างไม่สามารถทำงานให้แล้วเสร็จภายในกำหนดเวลา
หรือจะแล้วเสร็จล่าช้าเกินกว่ากำหนดเวลา หรือตกเป็นผู้ถูกพิทักษ์ทรัพย์เด็ดขาดหรือตกเป็นบุคคลล้มละลาย
หรือเพิกเฉยไม่ปฏิบัติตามคำสั่งของคณะกรรมการตรวจรับพัสดุ ผู้ว่าจ้างมีสิทธิที่จะบอกเลิกบันทึกข้อตกลงนี้ได้
และมีสิทธิจ้างผู้รับจ้างรายใหม่เข้าทำงานของผู้รับจ้างให้ลุล่วงไปได้ด้วย
การใช้สิทธิบอกเลิกบันทึกข้อตกลงนั้นไม่กระทบสิทธิของ
ผู้ว่าจ้างที่จะเรียกร้องค่าเสียหายจากผู้รับจ้าง</p>
    <p class='t-16 tab4'>การที่ผู้ว่าจ้างไม่ใช้สิทธิบอกเลิกบันทึกข้อตกลงดังกล่าวข้างต้นนั้น
ไม่เป็นเหตุให้ผู้รับจ้างพ้นจากความรับผิดตามบันทึกข้อตกลง</p>

<P class='t-16 tab3'><b>ข้อ ๕.การจ้างช่วง</b></P>
    <P class='t-16 tab4'>ผู้รับจ้างจะต้องไม่เอางานทั้งหมดหรือแต่บางส่วนของบันทึกข้อตกลงนี้ไปให้ผู้อื่นรับจ้างช่วงอีกทอดหนึ่ง
เว้นแต่การจ้างช่วงงานแต่บางส่วนที่ได้รับอนุญาตเป็นหนังสือจากผู้ว่าจ้างแล้ว การที่ผู้ว่าจ้างได้อนุญาตให้จ้างช่วงงานแต่บางส่วน 
ดังกล่าวนั้น ไม่เป็นเหตุให้ผู้รับจ้างหลุดพ้นจากความรับผิดหรือพันธะหน้าที่ตามบันทึกข้อตกลงนี้และผู้รับจ้างจะยังคงต้อง 
รับผิดในความผิดและความประมาทเลินเล่อของผู้รับจ้างช่วงหรือของตัวแทนหรือลูกจ้างของผู้รับจ้างช่วงนั้นทุกประการ
</P>
<P class='t-16 tab4'>กรณีผู้รับจ้างไปจ้างช่วงงานแต่บางส่วนโดยฝ่าฝืนความในวรรคหนึ่งผู้รับจ้างต้องชำระค่าปรับให้แก่ผู้ว่าจ้างเป็น 
จำนวนเงินในอัตราร้อยละ ๑๐.๐๐ (สิบ) ของวงเงินของงานที่จ้างช่วงตามบันทึกข้อตกลง ทั้งนี้ ไม่ตัดสิทธิผู้ว่าจ้างในการบอกเลิก บันทึกข้อตกลง</P>

<P class='t-16 tab3'><b>ข้อ ๖.ความรับผิดของผู้รับจ้าง</b></P>
<P class='t-16 tab4'>ผู้รับจ้างจะต้องรับผิดต่ออุบัติเหตุ ความเสียหาย หรือภยันตรายใดๆ อันเกิดจากการปฏิบัติงานของผู้รับจ้าง
และจะต้องรับผิดต่อความเสียหายจากการกระทำของลูกจ้างหรือตัวแทนของผู้รับจ้าง และจากการปฏิบัติงานของผู้รับจ้างช่วงด้วย (ถ้ามี)</P>
 <P class='t-16 tab4'>ความเสียหายใดๆ อันเกิดแก่งานที่ผู้รับจ้างได้ทำขึ้น แม้จะเกิดขึ้นเพราะเหตุสุดวิสัยก็ตาม 
ผู้รับจ้างจะต้องรับผิดชอบโดยซ่อมแซมให้คืนดีหรือเปลี่ยนให้ใหม่โดยค่าใช้จ่ายของผู้รับจ้างเอง เว้นแต่ความเสียหายนั้น เกิดจากความผิดของผู้ว่าจ้าง
ทั้งนี้ ความรับผิดของผู้รับจ้างดังกล่าวในข้อนี้จะสิ้นสุดลงเมื่อผู้ว่าจ้างได้รับมอบงานครั้งสุดท้าย</P>
<P class='t-16 tab4'>ผู้รับจ้างจะต้องรับผิดต่อบุคคลภายนอกในความเสียหายใดๆ อันเกิดจากการปฏิบัติงานของผู้รับจ้าง หรือลูกจ้าง 
หรือตัวแทนของผู้รับจ้าง รวมถึงผู้รับจ้างช่วง (ถ้ามี) ตามบันทึกข้อตกลงนี้ หากผู้ว่าจ้างถูกเรียกร้องหรือฟ้องร้อง 
หรือต้องชดใช้ค่าเสียหายให้แก่บุคคลภายนอกไปแล้ว ผู้รับจ้างจะต้องดำเนินการใดๆ เพื่อให้มีการว่าต่างแก้ต่างให้แก่ผู้ว่าจ้าง 
โดยค่าใช้จ่ายของผู้รับจ้างเอง รวมทั้งผู้รับจ้างจะต้องชดใช้ค่าเสียหายนั้นๆ ตลอดจนค่าใช้จ่ายใดๆ อันเกิดจากการถูกเรียกร้อง หรือถูกฟ้องร้องให้แก่ผู้ว่าจ้างทันที
</P>

  <P class='t-16 tab3'><b>ข้อ ๗.	การตรวจรับงานจ้าง</b></P>
    <P class='t-16 tab4'> เมื่อผู้ว่าจ้างได้ตรวจรับงานจ้างที่ส่งมอบและเห็นว่าถูกต้องครบถ้วนตามบันทึกข้อตกลงแล้ว
ผู้ว่าจ้างจะออกหลักฐานการรับมอบเป็นหนังสือไว้ให้เพื่อผู้รับจ้างนำมาเป็นหลักฐานประกอบการขอรับเงินค่างานจ้างนั้น</P>
  <P class='t-16 tab4'>ถ้าผลของการตรวจรับงานจ้างปรากฏว่างานจ้างที่ผู้รับจ้างส่งมอบไม่ตรงตามบันทึกข้อตกลงผู้ว่าจ้างทรง 
ไว้ซึ่งสิทธิที่จะไม่รับงานจ้างนั้นในกรณีเช่นว่านี้ ผู้รับจ้างต้องทำการแก้ไขให้ถูกต้องตามบันทึกข้อตกลงด้วยค่าใช้จ่ายของผู้รับจ้างเอง
และระยะเวลาที่เสียไปเพราะเหตุดังกล่าวผู้รับจ้างจะนำมาอ้างเป็นเหตุขอขยายเวลาส่งมอบงานจ้างตามบันทึกข้อตกลงหรือของด หรือลดค่าปรับไม่ได้</P>

   <P class='t-16 tab3'><b>ข้อ ๘.	รายละเอียดของงานจ้างคลาดเคลื่อน</b></P>
    <P class='t-16 tab4'>ผู้รับจ้างรับรองว่าได้ตรวจสอบและทำความเข้าใจในรายละเอียดของงานจ้างโดยถี่ถ้วนแล้ว 
หากปรากฏว่ารายละเอียดของงานจ้างนั้นผิดพลาดหรือคลาดเคลื่อนไปจากหลักการทางวิศวกรรมหรือทางเทคนิค 
ผู้รับจ้างตกลงที่จะปฏิบัติตามคำวินิจฉัยของผู้ว่าจ้าง คณะกรรมการตรวจรับพัสดุ เพื่อให้งานแล้วเสร็จบริบูรณ์ คำวินิจฉัยดังกล่าวให้ถือเป็นที่สุด
โดยผู้รับจ้างจะคิดค่าจ้าง ค่าเสียหาย หรือค่าใช้จ่ายใดๆ เพิ่มขึ้นจากผู้ว่าจ้าง หรือขอขยายอายุบันทึกข้อตกลงไม่ได้</P>
  
  <P class='t-16 tab3'><b>ข้อ ๙.	ค่าปรับ</b></P>
    <P class='t-16 tab4'>หากผู้รับจ้างไม่สามารถทำงานให้แล้วเสร็จภายในเวลาที่กำหนดไว้ในบันทึกข้อตกลง
และผู้ว่าจ้างยังมิได้บอกเลิกบันทึกข้อตกลง ผู้รับจ้างจะต้องชำระค่าปรับให้แก่ผู้ว่าจ้างเป็นจำนวนเงิน
วันละ ........... บาท (...........) นับถัดจากวันที่ครบกำหนดเวลาแล้วเสร็จของงานตามบันทึกข้อตกลง 
หรือวันที่ผู้ว่าจ้างได้ขยายเวลาทำงานให้จนถึงวันที่ทำงานแล้วเสร็จจริง นอกจากนี้ ผู้รับจ้างยอมให้ผู้ว่าจ้างเรียกค่าเสียหาย 
อันเกิดขึ้นจากการที่ผู้รับจ้างทำงานล่าช้าเฉพาะส่วนที่เกินกว่าจำนวนค่าปรับดังกล่าวได้อีกด้วย</P>
     <P class='t-16 tab4'>ในระหว่างที่ผู้ว่าจ้างยังมิได้บอกเลิกบันทึกข้อตกลงนั้น หากผู้ว่าจ้างเห็นว่าผู้รับจ้างจะไม่สามารถ 
ปฏิบัติตามบันทึกข้อตกลงต่อไปได้ ผู้ว่าจ้างจะใช้สิทธิบอกเลิกบันทึกข้อตกลงและใช้สิทธิตามข้อ ๑๐ ก็ได้และถ้าผู้ว่าจ้าง 
ได้แจ้งข้อเรียกร้องไปยังผู้รับจ้างเมื่อครบกำหนดเวลาแล้วเสร็จของงานขอให้ชำระค่าปรับแล้วผู้ว่าจ้างมีสิทธิที่จะปรับผู้รับจ้างจน ถึงวันบอกเลิกบันทึกข้อตกลงได้อีกด้วย</P>

<P class='t-16 tab3'><b>ข้อ ๑๐.สิทธิของผู้ว่าจ้างภายหลังบอกเลิกบันทึกข้อตกลง
</b></P>
    <P class='t-16 tab4'>ในกรณีที่ผู้ว่าจ้างบอกเลิกบันทึกข้อตกลง ผู้ว่าจ้างอาจทำงานนั้นเองหรือว่าจ้างผู้อื่นให้ทำงานนั้นต่อจน  
แล้วเสร็จก็ได้และในกรณีดังกล่าว ผู้รับจ้างจะต้องรับผิดชอบในค่าเสียหายที่เกิดขึ้นรวมทั้งค่าใช้จ่ายที่เพิ่มขึ้นในการทำงานนั้นต่อ 
ให้แล้วเสร็จตามบันทึกข้อตกลง และผู้ว่าจ้างจะหักเอาจากจำนวนเงินใดๆ ที่จะจ่ายให้แก่ผู้รับจ้างก็ได้</P>

      <P class='t-16 tab3'><b>ข้อ ๑๑. อื่นๆ </b></P>
     <P class='t-16 tab4'>การจ้างเหมาบริการตามบันทึกข้อตกลงนี้ ไม่ทำให้ผู้รับจ้างมีฐานะเป็นลูกจ้างของผู้ว่าจ้างหรือมีความสัมพันธ์ 
ในฐานะเป็นลูกจ้างตามกฎหมายแรงงาน หรือกฎหมายว่าด้วยประกันสังคม
</P>
    <P class='t-16 tab4'>บันทึกข้อตกลงนี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน ทั้งสองฝ่ายได้อ่านและเข้าใจข้อความ
โดยละเอียดตลอดแล้ว จึงได้ลงลายมือชื่อไว้เป็นสำคัญต่อหน้าพยาน และได้เก็บบันทึกข้อตกลงนี้ไว้ฝ่ายละหนึ่งฉบับ</P>

  
</br>
</br>
<!-- 🔹 รายชื่อผู้ลงนาม -->
{signatoryWithLogoHtml.ToString()}

</div>
</body>
</html>
";

       
        return html;
    }
}
