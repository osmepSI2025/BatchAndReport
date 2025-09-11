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
using Spire.Doc.Documents;
using System.Text;
using System.Threading.Tasks;
using static SkiaSharp.HarfBuzz.SKShaper;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
public class WordEContract_JointOperationService
{
    private readonly WordServiceSetting _w;
    private readonly E_ContractReportDAO _eContractReportDAO;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter

    public WordEContract_JointOperationService(
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
    public async Task<byte[]> OnGetWordContact_JointOperationService(string conId)
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

    public async Task<string> OnGetWordContact_JointOperationServiceHtmlToPDF(string conId)
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
        string strAttorneyLetterDate_CP = CommonDAO.ToArabicDateStringCovert(dataResult.CP_S_AttorneyLetterDate ?? DateTime.Now);
        string strAttorneyOsmep = "";
        var HtmlAttorneyOsmep = new StringBuilder();
        if (dataResult.AttorneyFlag == true)
        {
            strAttorneyOsmep = "ผู้มีอำนาจกระทำการแทนปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ เลขคำสั่งที่ " + dataResult.AttorneyLetterNumber + " ฉบับลงวันที่ " + strAttorneyLetterDate + "";

        }
        else
        {
            strAttorneyOsmep = "";
        }
        string strAttorney = "";
        var HtmlAttorney = new StringBuilder();
        if (dataResult.CP_S_AttorneyFlag == true)
        {
            strAttorney = "ผู้มีอำนาจ กระทำการแทน ปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ ฉบับลงวันที่ " + strAttorneyLetterDate_CP + "";

        }
        else
        {
            strAttorney = "";
        }
        #endregion

        var strDateTH = CommonDAO.ToThaiDateString(dataResult.Contract_SignDate ?? DateTime.Now);
        var purposeList = await _eContractReportDAO.GetJOAPoposeAsync(conId);

        #region signlist joa
        // call function RenderSignatory
        var signatoryTableHtml = "";
        if (dataResult.Signatories.Count > 0)
        {
            signatoryTableHtml = await _eContractReportDAO.RenderSignatory(dataResult.Signatories);
        
        }


        #endregion signlist



        // Use signatoryTableHtml in your final HTML output


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
         body, table, th, td, span, strong, i, u {{
            font-size: 22px;
           font-family: 'THSarabunNew', Arial, sans-serif !important;
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
    <div class='t-22 text-center'>สัญญาร่วมดำเนินการ</div>
    <div class='t-22 text-center'><b>โครงการ {dataResult.Project_Name}</b></div>
    <div class='t-16 text-center'><b>ระหว่าง</b></div>
    <div class='t-18 text-center'><b>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</b></div>
    <div class='t-18 text-center'><b>กับ</b></div>
    <div class='t-18 text-center'><b>{dataResult.Organization ?? ""}</b></div>
</br>
    <P class='t-16 tab3'>
        สัญญาร่วมดำเนินการฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เมื่อวันที่ {strDateTH[0]} เดือน {strDateTH[1]} พ.ศ.{strDateTH[2]} ระหว่าง
    </P>
    <P class='t-16 tab3'>
    <B>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B>  โดย {dataResult.OSMEP_NAME} ตำแหน่ง {dataResult.OSMEP_POSITION} {strAttorneyOsmep} สำนักงานตั้งอยู่เลขที่ 120 หมู่ 3 ศูนย์ราชการเฉลิมพระเกียรติ 80 พรรษา 5 ธันวาคม 2550. (อาคารซี) ชั้น 2, 10, 11 ถนนแจ้งวัฒนะ แขวงทุ่งสองห้อง เขตหลักสี่ กรุงเทพ 10210 ซึ่งต่อไป ในสัญญาฉบับนี้จะเรียกว่า“สสว.” ฝ่ายหนึ่ง กับ
    </P>
    <P class='t-16 tab3'>
        {dataResult.Organization ?? ""} โดย {dataResult.CP_S_NAME} ตำแหน่ง {dataResult.CP_S_POSITION} {strAttorney} ซึ่งต่อไปในสัญญาฉบับนี้จะเรียกว่า “{dataResult.Organization}” อีกฝ่ายหนึ่ง
    </P>
    <P class='t-16 tab0'><B>วัตถุประสงค์ตามสัญญาร่วมดำเนินการ</B></P>
    <P class='t-16 tab3'>คู่สัญญาทั้งสองฝ่ายมีความประสงค์ที่จะร่วมมือกันเพื่อดำเนินการภายใต้โครงการ {dataResult.Project_Name} ซึ่งต่อไปในสัญญานี้จะเรียกว่า “โครงการ” โดยมีรายละเอียดโครงการ แผนการดำเนินงาน
 แผนการใช้จ่ายเงิน (และอื่นๆ เช่น คู่มือดำเนินโครงการ) และบรรดาเอกสาร แนบท้ายสัญญาฉบับนี้ ซึ่งให้ถือเป็นส่วนหนึ่งของสัญญาฉบับนี้ มีระยะเวลาตั้งแต่วันที่ {CommonDAO.ToThaiDateStringCovert(dataResult.Contract_Start_Date ?? DateTime.Now)} จนถึงวันที่ {CommonDAO.ToThaiDateStringCovert(dataResult.Contract_End_Date ?? DateTime.Now)} โดยมีวัตถุประสงค์ในการดำเนินโครงการ ดังนี้
    </P>
{(purposeList != null && purposeList.Count != 0
    ? $"<div class='t-16 tab2'>{string.Join("<br/>", purposeList.Select(p => p.Detail))}</div>"
    : "")}    

<P class='t-16 tab3'><B>ข้อ 1 ขอบเขตหน้าที่ของ “สสว.”</B></P>
    <P class='t-16 tab4'>1.1 ตกลงร่วมดำเนินการโครงการโดยสนับสนุนงบประมาณ จำนวน {dataResult.Contract_Value?.ToString("N2") ?? "0.00"} บาท ( {CommonDAO.NumberToThaiText(dataResult.Contract_Value ?? 0)} ) ซึ่งได้รวมภาษีมูลค่าเพิ่ม ตลอดจนค่าภาษีอากรอื่นๆ แล้วให้กับ “{dataResult.Organization}” และการใช้จ่ายเงินให้เป็นไปตามแผนการจ่ายเงินตามเอกสารแนบท้ายสัญญา</P>
    <P class='t-16 tab4'>1.2 ประสานการดำเนินโครงการ เพื่อให้บรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์</P>
    <P class='t-16 tab4'>1.3 กำกับ ติดตามและประเมินผลการดำเนินงานของโครงการ</P>
    <P class='t-16 tab3'><B>ข้อ 2 ขอบเขตหน้าที่ของ “{dataResult.Organization}”</B></P>
    <P class='t-16 tab4'>2.1 ตกลงที่จะร่วมดำเนินการโครงการตามวัตถุประสงค์ของการโครงการและขอบเขตการ ดำเนินการ ตามรายละเอียดโครงการ แผนการดำเนินการ และแผนการใช้จ่ายเงิน (และอื่นๆ เช่น คู่มือดำเนิน โครงการ) ที่แนบท้าย สัญญาฉบับนี้</P>
    <P class='t-16 tab4'>2.2 ต้องดำเนินโครงการ ปฏิบัติตามแผนการดำเนินงาน แผนการใช้จ่ายเงิน (หรืออาจมีคู่มือการ ดำเนินโครงการก็ได้) อย่างเคร่งครัดและ ให้แล้วเสร็จภายในระยะเวลาโครงการ หากไม่ดำเนินโครงการให้แล้วเสร็จ ตามที่กำหนดยินยอม ชำระค่าปรับให้แก่ สสว. ในอัตราร้อยละ 0.1 ของจำนวนงบประมาณ ที่ได้รับการสนับสนุนทั้งหมดต่อวัน นับถัดจากวันที่กำหนด แล้วเสร็จ และถ้าหากเห็นว่า “{dataResult.Organization}” ไม่อาจปฏิบัติตามสัญญาต่อไปได้ “{dataResult.Organization}” ยินยอมให้ สสว.ใช้สิทธิบอกเลิกสัญญาได้ทันที</P>
    <P class='t-16 tab4'>2.3 ต้องประสานการดำเนินโครงการ เพื่อให้บรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์</P>
    <P class='t-16 tab4'>2.4 ต้องให้ความร่วมมือกับ สสว. ในการกำกับ ติดตามและประเมินผลการดำเนินงาน ของโครงการ</P>
    <P class='t-16 tab3'><B>ข้อ 3 อื่น ๆ</B></P>
    <div class='t-16 tab4'>3.1 หากคู่สัญญาฝ่ายใดฝ่ายหนึ่งประสงค์จะขอแก้ไข เปลี่ยนแปลง ขยายระยะเวลาของ โครงการ จะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษร และต้องได้รับความยินยอม เป็นลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และต้องทำเอกสารแก้ไข เปลี่ยนแปลง ขยายระยะเวลา เพื่อลงนามยินยอม ทั้งสองฝ่าย</div>
    <P class='t-16 tab4'>3.2 หากคู่สัญญาฝ่ายใดฝ่ายหนึ่งประสงค์จะขอบอกเลิกสัญญาก่อนครบกำหนดระยะเวลา ดำเนินโครงการ จะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษรไม่น้อยกว่า 30 วัน และต้องได้รับความยินยอมเป็นลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และ “{dataResult.Organization}” จะต้องคืนเงินในส่วนที่ยังไม่ได้ใช้จ่ายหรือส่วนที่เหลือทั้งหมดพร้อมดอกผล (ถ้ามี) ให้แก่ สสว. ภายใน 15 วัน นับจากวันที่ได้รับหนังสือของฝ่ายที่ยินยอมให้บอกเลิก</P>
    <P class='t-16 tab4'>3.3 สสว. อาจบอกเลิกสัญญาได้ทันที หากตรวจสอบ หรือปรากฏข้อเท็จจริงว่า การใช้จ่ายเงินของ “{dataResult.Organization}” ไม่เป็นไปตามวัตถุประสงค์ของโครงการ แผนการดำเนินงาน และแผนการใช้จ่ายเงิน (และอื่นๆ เช่น คู่มือดำเนินโครงการ) ทั้งมีสิทธิเรียกเงินคงเหลือคืนทั้งหมดพร้อมดอกผล (ถ้ามี) ได้ทันที</P>
    <P class='t-16 tab4'>3.4 ทรัพย์สินใดๆ และ/หรือ สิทธิใดๆ ที่ได้มาจากเงินสนับสนุนตามสัญญาร่วมดำเนินการ ฉบับนี้ เมื่อสิ้นสุดโครงการให้ตกได้แก่ สสว. ทั้งสิ้น เว้นแต่ สสว. จะกำหนดให้เป็นอย่างอื่น</P>
    <p class='t-16 tab4'>3.5 “{dataResult.Organization}” ต้องไม่ดำเนินการในลักษณะการจ้างเหมา กับหน่วยงาน องค์กร หรือบุคคลอื่นๆ ยกเว้นกรณีการจัดหา จัดจ้าง เป็นกิจกรรมหรือเป็นเรื่อง ๆ</p>
    <p class='t-16 tab4'>3.6 ในกรณีที่การดำเนินการตามสัญญาฉบับนี้ เกี่ยวข้องกับข้อมูลส่วนบุคคล และการคุ้ม ครองทรัพย์สินทางปัญญา “{dataResult.Organization}” จะต้องปฏิบัติตามกฎหมายว่าด้วยการคุ้มครอง ข้อมูลส่วนบุคคลและการคุ้มครองทรัพย์สินทางปัญญาอย่างเคร่งครัดและหากเกิดความเสียหายหรือมีการฟ้องร้องใดๆ “{dataResult.Organization}” จะต้องเป็นผู้รับผิดชอบต่อการละเมิดบทบัญญัติแห่งกฎหมายดังกล่าวแต่เพียงฝ่ายเดียวโดยสิ้นเชิง</p>
    <P class='t-16 tab3'>สัญญาฉบับนี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน ทั้งสองฝ่ายได้อ่านและเข้าใจข้อ ความโดยละเอียดแล้ว จึงได้ลงลายมือชื่อพร้อมประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยานและยึดถือ ไว้ฝ่ายละฉบับ</P>
 
</br>
</br>
<!-- 🔹 รายชื่อผู้ลงนาม -->


{signatoryTableHtml}

</div>
</body>
</html>
";

 

        return html;
    }
}
