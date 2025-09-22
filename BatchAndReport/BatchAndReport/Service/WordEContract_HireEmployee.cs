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
    public async Task<byte[]> OnGetWordContact_HireEmployee(string id)
    {
        try {
            var result = await _e.GetECAsync(id);
            var stream = new MemoryStream();
            using (var wordDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                // Styles
                var stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylePart.Styles = WordServiceSetting.CreateDefaultStyles();
                var body = mainPart.Document.AppendChild(new Body());

                // --- Logo section: large, centered, with whitespace above and below ---
                var imagePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
                if (System.IO.File.Exists(imagePath))
                {
                    // Add empty paragraph above logo for spacing
                    //  body.AppendChild(EmptyParagraph());

                    var imagePart = mainPart.AddImagePart(ImagePartType.Png);
                    using (var imgStream = new FileStream(imagePath, FileMode.Open))
                    {
                        imagePart.FeedData(imgStream);
                    }
                    // Make logo larger (e.g., 240x80 px)
                    var element = WordServiceSetting.CreateImage(mainPart.GetIdOfPart(imagePart), 240, 80);
                    var logoPara = new Paragraph(
                        new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                        element
                    );
                    body.AppendChild(logoPara);
                }
                // 2. Document title and subtitle

                body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("สัญญาจ้างลูกจ้าง", "000000", "36"));
                string strcontractsign = CommonDAO.ToArabicDateStringCovert(result.ContractSignDate ?? DateTime.Now);
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สัญญาฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เลขที่ 21 ถนนวิภาวดีรังสิต เขตจตุจักร กรุงเทพมหานคร เมื่อ" + strcontractsign + "", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ระหว่าง สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม โดย........................................." +
                    "ผู้อำนวยการฝ่ายศูนย์ให้บริการ SMEs ครบวงจร สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ผู้รับมอบหมายตามคำสั่งสำนักงานฯ ที่ 629/2564 ลงวันที่ 30 กันยายน 2564 ซึ่งต่อไปในสัญญานี้จะเรียกว่า “ผู้ว่าจ้าง”", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ฝ่ายหนึ่ง กับ " + result.SignatoryName + " เลขประจำตัวประชาชน " + result.IdenID + " อยู่บ้านเลขที่ " + result.EmpAddress + " " +
                    "ซึ่งต่อไปในสัญญานี้จะเรียกว่า “ลูกจ้าง” อีกฝ่ายหนึ่ง โดยทั้งสองฝ่ายได้ตกลงทำร่วมกัน</br>ดังมีรายละเอียดต่อไปนี้", null, "32"));

                string strHiringStart = CommonDAO.ToArabicDateStringCovert(result.HiringStartDate ?? DateTime.Now);
                string strHiringEnd = CommonDAO.ToArabicDateStringCovert(result.HiringEndDate ?? DateTime.Now);
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("1.ผู้ว่าจ้างตกลงจ้างลูกจ้างปฏิบัติงานกับผู้ว่าจ้าง โดยให้ปฏิบัติงานภายใต้งาน " + result.WorkDetail + "  ในตำแหน่ง " + result.WorkPosition + " ปฏิบัติหน้าที่ ณ ศูนย์กลุ่มจังหวัดให้บริการ SME ครบวงจร ..... " +
                    "โดยมีรายละเอียดหน้าที่ความรับผิดชอบปรากฏตามเอกสารแนบท้ายสัญญาจ้าง ตั้งแต่" + strHiringStart + " ถึง" + strHiringEnd + "", null, "32"));

                string strSalary = CommonDAO.NumberToThaiText(result.Salary ?? 0);
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("2.ผู้ว่าจ้างจะจ่ายค่าจ้างให้แก่ลูกจ้างในระหว่างระยะเวลาการปฏิบัติงานของลูกจ้างตามสัญญานี้ในอัตราเดือนละ " + result.Salary + "บาท (" + strSalary + ")" +
                    "โดยจะจ่ายให้ในวันทำการก่อนวันทำการสุดท้ายของธนาคารในเดือนนั้นสามวันทำการ และนำเข้าบัญชีเงินฝากของลูกจ้าง ณ ที่ทำการของผู้ว่าจ้าง หรือ ณ ที่อื่นใดตามที่ผู้ว่าจ้างกำหนด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.ในการจ่ายค่าจ้าง และ/หรือ เงินในลักษณะอื่นให้แก่ลูกจ้าง ลูกจ้างตกลงยินยอมให้ผู้ว่าจ้างหักภาษี ณ ที่จ่าย และ/หรือ เงินอื่นใดที่ต้องหักโดยชอบด้วยระเบียบ ข้อบังคับของผู้ว่าจ้างหรือตามกฎหมายที่เกี่ยวข้อง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("4.ตลอดระยะเวลาการปฏิบัติงานตามสัญญานี้ ลูกจ้างมีสิทธิได้รับสิทธิประโยชน์อื่น ๆ ตามที่กำหนดไว้ใน ระเบียบ ข้อบังคับ คำสั่ง หรือประกาศใด ๆ ตามที่ผู้ว่าจ้างกำหนด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("5.ผู้ว่าจ้างจะทำการประเมินผลการปฏิบัติงานอย่างน้อยปีละสองครั้ง ตามหลักเกณฑ์และวิธีการที่ผู้ว่าจ้างกำหนด ทั้งนี้ หากผลการประเมินไม่ผ่านตามหลักเกณฑ์ที่กำหนด ผู้ว่าจ้างมีสิทธิบอกเลิกสัญญาจ้างได้ และลูกจ้างไม่มีสิทธิเรียกร้องเงินชดเชยหรือเงินอื่นใด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("6.ตลอดระยะเวลาการปฏิบัติงานตามสัญญานี้ ลูกจ้างจะต้องปฏิบัติตามกฎ ระเบียบ ข้อบังคับ คำสั่งหรือประกาศใด ๆ ของผู้ว่าจ้าง " +
                    "ตลอดจนมีหน้าที่ต้องรักษาวินัยและยอมรับการลงโทษทางวินัยของผู้ว่าจ้างโดยเคร่งครัด และยินยอมให้ถือว่า กฎหมาย ระเบียบ ข้อบังคับ หรือคำสั่งต่าง ๆ ของผู้ว่าจ้างเป็นส่วนหนึ่งของสัญญาจ้างนี้", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ในกรณีลูกจ้างจงใจขัดคำสั่งโดยชอบของผู้ว่าจ้างหรือละเลยไม่นำพาต่อคำสั่งเช่นว่านั้นเป็นอาจิณ หรือประการอื่นใด อันไม่สมควรกับการปฏิบัติหน้าที่ของตนให้ลุล่วงไปโดยสุจริตและถูกต้อง ลูกจ้างยินยอมให้ผู้ว่าจ้างบอกเลิกสัญญาจ้างโดยมิต้องบอกกล่าวล่วงหน้า", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("7. ลูกจ้างต้องปฏิบัติงานให้กับผู้ว่าจ้าง ตามที่ได้รับมอบหมายด้วยความซื่อสัตย์ สุจริต <br>และตั้งใจปฏิบัติงานอย่างเต็มกำลังความสามารถของตน โดยแสวงหาความรู้และทักษะเพิ่มเติมหรือกระทำการใด " +
                    "เพื่อให้ผลงานในหน้าที่มีคุณภาพดีขึ้น ทั้งนี้ ต้องรักษาผลประโยชน์และชื่อเสียงของผู้ว่าจ้าง และไม่เปิดเผยความลับหรือข้อมูลของทางราชการให้ผู้หนึ่งผู้ใดทราบ โดยมิได้รับอนุญาตจากผู้รับผิดชอบงานนั้น ๆ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("8. สัญญานี้สิ้นสุดลงเมื่อเข้ากรณีใดกรณีหนึ่ง ดังต่อไปนี้", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("8.1 สิ้นสุดระยะเวลาตามสัญญาจ้าง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("8.2 เมื่อผู้ว่าจ้างบอกเลิกสัญญาจ้าง หรือลูกจ้างบอกเลิกสัญญาจ้างตามข้อ 10", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("8.3 ลูกจ้างกระทำการผิดวินัยร้ายแรง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("8.4 ลูกจ้างไม่ผ่านการประเมินผลการปฏิบัติงานของลูกจ้างตามข้อ 5", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("9. ในกรณีที่สัญญาสิ้นสุดตามข้อ 8.3 และ 8.4 ลูกจ้างยินยอมให้ผู้ว่าจ้างสั่งให้ลูกจ้างพ้นสภาพการเป็นลูกจ้างได้ทันที โดยไม่จำเป็นต้องมีหนังสือว่ากล่าวตักเตือน และผู้ว่าจ้างไม่ต้องจ่ายค่าชดเชยหรือเงินอื่นใดให้แก่ลูกจ้างทั้งสิ้น เว้นแต่ค่าจ้างที่ลูกจ้างจะพึงได้รับตามสิทธิ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("10. ลูกจ้างมีสิทธิบอกเลิกสัญญาจ้างได้ก่อนสัญญาครบกำหนด โดยทำหนังสือแจ้งเป็นลายลักษณ์อักษรต่อผู้ว่าจ้างได้ทราบล่วงหน้าไม่น้อยกว่า 30 วัน เมื่อผู้ว่าจ้างได้อนุมัติแล้ว ให้ถือว่าสัญญาจ้างนี้ได้สิ้นสุดลง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("11. ในกรณีที่ลูกจ้างกระทำการใดอันทำให้ผู้ว่าจ้างได้รับความเสียหาย ไม่ว่าเหตุนั้นผู้ว่าจ้างจะนำมาเป็นเหตุบอกเลิกสัญญาจ้างหรือไม่ก็ตาม ผู้ว่าจ้างมีสิทธิจะเรียกร้องค่าเสียหาย และลูกจ้างยินยอมชดใช้ค่าเสียหายตามที่ผู้ว่าจ้างเรียกร้องทุกประการ ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("12. ลูกจ้างจะต้องไม่เปิดเผยหรือบอกกล่าวอัตราค่าจ้างของลูกจ้างให้แก่บุคคลใดทราบ ไม่ว่าจะโดยวิธีใดหรือเวลาใด เว้นแต่จะเป็นการกระทำตามกฎหมายหรือคำสั่งศาล", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สัญญาฉบับนี้ได้จัดทำขึ้นเป็นสัญญาอิเล็กทรอนิกส์คู่สัญญาได้อ่านตรวจสอบและทำความเข้าใจข้อความในสัญญาฉบับนี้โดยละเอียดแล้ว จึงได้ลงลายมือชื่ออิเล็กทรอนิกส์ไว้เป็นหลักฐาน ณ วัน เดือน ปี ดังกล่าวข้างต้น และมีพยานรู้ถึงการลงนามของคู่สัญญา และคู่สัญญาต่างฝ่ายต่างเก็บรักษาไฟล์สัญญาอิเล็กทรอนิกส์ฉบับนี้ไว้เป็นหลักฐาน", null, "32"));

                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // --- Signature Table Section ---

                // Create a table for signatures
                var signatureTable = new Table(
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
                    // Row 1: Employer / Employee
                    new TableRow(
                        new TableCell(
                            WordServiceSetting.CenteredParagraph("ลงชื่อ__________________________ผู้ว่าจ้าง")
                        ),
                        new TableCell(
                            WordServiceSetting.CenteredParagraph("ลงชื่อ__________________________ลูกจ้าง")
                        )
                    ),
                    // Row 2: ( ) / ( )
                    new TableRow(
                        new TableCell(
                            WordServiceSetting.CenteredParagraph("(__________________________)")
                        ),
                        new TableCell(
                            WordServiceSetting.CenteredParagraph("(__________________________)")
                        )
                    ),
                    // Row 3: Witness / Witness
                    new TableRow(
                        new TableCell(
                            WordServiceSetting.CenteredParagraph("ลงชื่อ__________________________พยาน")
                        ),
                        new TableCell(
                            WordServiceSetting.CenteredParagraph("ลงชื่อ__________________________พยาน")
                        )
                    ),
                    // Row 4: ( ) / ( )
                    new TableRow(
                        new TableCell(
                            WordServiceSetting.CenteredParagraph("(__________________________)")
                        ),
                        new TableCell(
                            WordServiceSetting.CenteredParagraph("(__________________________)")
                        )
                    )
                );

                // Add the signature table to the document
                body.AppendChild(signatureTable);
                body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("เอกสารแนบท้ายสัญญาจ้างลูกจ้าง", "36"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("งานศูนย์ให้บริการ SMEs ครบวงจร", "36"));


                body.AppendChild(WordServiceSetting.NormalParagraph("หน้าที่ความรับผิดชอบ : เจ้าหน้าที่ศูนย์กลุ่มจังหวัดให้บริการ SMEs ครบวงจร และ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraph("                เจ้าหน้าที่ศูนย์ให้บริการ SMEs ครบวงจร กรุงเทพมหานคร", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_1Tabs("- การปรับปรุงข้อมูลผู้ประกอบการ SME (ไม่น้อยกว่า 30 ราย/เดือน)", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_1Tabs("- การให้บริการคำปรึกษา แนะนำทางธุรกิจ อาทิเช่น ด้านบัญชี การเงิน การตลาด </br>การบริหารจัดการ การผลิต กฎหมาย เทคโนโลยีสารสนเทศ และอื่น ๆ ที่เกี่ยวข้องทางธุรกิจ (ไม่น้อยกว่า 30 ราย/เดือน)", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_1Tabs("- สนับสนุน เสนอแนะแนวทางการแก้ไขปัญหาให้ SME ได้รับประโยชน์ตามมาตรการของภาครัฐ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_1Tabs("- สนับสนุนการพัฒนาเครือข่ายหน่วยงานให้บริการส่งเสริม SME ให้บริการส่งต่อภายใต้</br>หน่วยงานพันธมิตร การติดตามผลและประสานงานแก้ไขปัญหา", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_1Tabs("- สนับสนุนนโยบาย มาตรการ และการทำงานของ สสว. ในการสร้าง ประสาน เชื่อมโยง</br>เครือข่ายในพื้นที่ (รูปแบบ Online & Offline) เพื่อสนับสนุนการปฏิบัติงานตามภารกิจ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_1Tabs("- สนับสนุนจัดทำข้อมูล SME จังหวัด เพื่อนำข้อมูลมาใช้ประโยชน์ในการเสนอแนะทางธุรกิจแก่ SME และเชื่อมโยงไปสู่การแก้ปัญหาหรือการจัดทำมาตรการภาครัฐ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_1Tabs("- ปฏิบัติงานภายใต้การบังคับบัญชาของผู้จัดการศูนย์กลุ่มจังหวัดฯ หรือ ผู้จัดการศูนย์ให้</br>บริการ SMEs ครบวงจร กรุงเทพมหานคร ตามประกาศ สสว. และเข้าร่วมกิจกรรมต่าง ๆ ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_1Tabs("- กำกับดูแลข้อมูลตาม พ.ร.บ.การคุ้มครองข้อมูลส่วนบุคคล", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_1Tabs("- งานอื่น ๆ ตามที่ได้รับมอบหมาย", null, "32"));

                // --- Add footer with page number centered ---
                var footerPart = mainPart.AddNewPart<FooterPart>();
                string footerPartId = mainPart.GetIdOfPart(footerPart);
                footerPart.Footer = new Footer(
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification() { Val = JustificationValues.Center }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                            ),
                            new FieldChar() { FieldCharType = FieldCharValues.Begin }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                            ),
                            new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                            ),
                            new FieldChar() { FieldCharType = FieldCharValues.Separate }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                            ),
                            new Text("1")
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                            ),
                            new FieldChar() { FieldCharType = FieldCharValues.End }
                        )
                    )
                );

                var sectionProps = new SectionProperties(

                    new FooterReference() { Type = HeaderFooterValues.Default, Id = footerPartId },
                    new DocumentFormat.OpenXml.Wordprocessing.PageSize() { Width = 11906, Height = 16838 }, // A4 size
                    new PageMargin() { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720, Gutter = 0 }
                );
                body.AppendChild(sectionProps);
            }
            stream.Position = 0;
            return stream.ToArray();
        }
        catch (Exception ex)
        {
            throw new Exception("Error in OnGetWordContact_HireEmployee: " + ex.Message, ex);
        }


    }
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
            var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg").Replace("\\", "/");

            // ── 3) ข้อความ/วันที่ (ป้องกัน null ด้วย ??) ────────────────────────────
            string strAttorneyLetterDate = CommonDAO.ToArabicDateStringCovert(result.AttorneyLetterDate ?? DateTime.Now);
            string strAttorney =
                result.AttorneyFlag == true
                ? $"ผู้รับมอบหมายตามคำสั่งสำนักงานฯ ที่ {result.AttorneyLetterNumber ?? ""} ลง {strAttorneyLetterDate}"
                : "";

            string strcontractsign = CommonDAO.ToArabicDateStringCovert(result.ContractSignDate ?? DateTime.Now);
            string strHiringStart = CommonDAO.ToArabicDateStringCovert(result.HiringStartDate ?? DateTime.Now);
            string strHiringEnd = CommonDAO.ToArabicDateStringCovert(result.HiringEndDate ?? DateTime.Now);
            string strSalary = CommonDAO.NumberToThaiText(result.Salary ?? 0);

            // ── 4) signers (กัน null) + render 2 คอลัมน์ + ตราประทับ 1 อัน ────────
            var signlist = await _eContractReportDAO.GetSignNameAsync(id, typeContact) ?? new List<E_ConReport_SignatoryModels?>();
            var safeList = signlist.Where(s => s != null).ToList();

            var roleByType = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["OSMEP_S"] = "ผู้ว่าจ้าง",
                ["OSMEP_W"] = "พยาน",
                ["CP_S"] = "ลูกจ้าง",
                ["CP_W"] = "พยาน"
            };

            var leftSigners = safeList
                .Where(s => {
                    var t = (s as dynamic)?.Signatory_Type as string;
                    return t == "OSMEP_S" || t == "OSMEP_W";
                })
                .OrderBy(s => ((s as dynamic)?.Signatory_Type as string) == "OSMEP_S" ? 0 : 1)
                .ToList();

            var rightSigners = safeList
                .Where(s => {
                    var t = (s as dynamic)?.Signatory_Type as string;
                    return t == "CP_S" || t == "CP_W";
                })
                .OrderBy(s => ((s as dynamic)?.Signatory_Type as string) == "CP_S" ? 0 : 1)
                .ToList();

            // ✅ ให้ตราประทับออก "แค่ 1 อัน" ทั้งเอกสาร
            bool sealAdded = false;

            string RenderSignatureBlock(dynamic signer, bool isCompanySide)
            {
                string signType = signer?.Signatory_Type as string ?? "";
                roleByType.TryGetValue(signType, out var roleLabel);
                roleLabel ??= "";

                // ลายเซ็น
                string signatureHtml;
                string ds = (string)(signer?.DS_FILE ?? "");
                if (!string.IsNullOrEmpty(ds) && ds.IndexOf("<content>", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    try
                    {
                        int s1 = ds.IndexOf("<content>", StringComparison.OrdinalIgnoreCase) + "<content>".Length;
                        int s2 = ds.IndexOf("</content>", StringComparison.OrdinalIgnoreCase);
                        string b64 = ds.Substring(s1, s2 - s1).Trim();
                        signatureHtml = "<div class='sign-img'>ลงชื่อ <img src='data:image/png;base64," + b64 + "' alt='signature' />" + roleLabel + "</div>";
                    }
                    catch
                    {
                        signatureHtml = "<div class='sign-line'>ลงชื่อ...................." + roleLabel + "</div>";
                    }
                }
                else
                {
                    signatureHtml = "<div class='sign-line'>ลงชื่อ...................." + roleLabel + "</div>";
                }

                string name = System.Net.WebUtility.HtmlEncode((string)(signer?.Signatory_Name ?? ""));
                string unit = System.Net.WebUtility.HtmlEncode((string)(signer?.BU_UNIT ?? ""));

                // ตราประทับ: แสดงเฉพาะครั้งแรกที่พบ CP_S ฝั่งขวา
                string sealBlock = "";
                if (!sealAdded && isCompanySide && string.Equals(signType, "CP_S", StringComparison.OrdinalIgnoreCase))
                {
                    string rawSeal = (string)(signer?.Company_Seal ?? "");
                    if (!string.IsNullOrEmpty(rawSeal) && rawSeal.IndexOf("<content>", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        try
                        {
                            int s1 = rawSeal.IndexOf("<content>", StringComparison.OrdinalIgnoreCase) + "<content>".Length;
                            int s2 = rawSeal.IndexOf("</content>", StringComparison.OrdinalIgnoreCase);
                            string b64 = rawSeal.Substring(s1, s2 - s1).Trim();

                            sealBlock =
    $@"<div class='seal-img'><img src='data:image/png;base64,{b64}' alt='company-seal' /></div>
<div class='seal-caption'>(ตราประทับ บริษัท)</div>";
                            sealAdded = true;
                        }
                        catch
                        {
                            // ถ้าแตกก็ไม่ใส่ seal และไม่ set flag จะปล่อยให้ signer ถัดไปลองได้
                        }
                    }
                }

                return
    $@"<div class='sign-block'>
    {signatureHtml}
    <div class='sign-name'>({name})</div>
    <div class='sign-unit'>{unit}</div>
    {sealBlock}
</div>";
            }
            

            string RenderFixedSignatureBlock()
            {
                var empName = string.IsNullOrWhiteSpace(result?.EmploymentName)
                ? "........................................."
                : System.Net.WebUtility.HtmlEncode(result.EmploymentName);

                return $@"
<div class='sign-block' style='margin-top:60px;'>
    <div class='sign-line'> </div>
    <div class='sign-line'>ลงชื่อ_______________________ลูกจ้าง</div>
    <div class='sign-name'>({empName})</div>
</div>
<div class='sign-block' style='margin-top:60px;'>
    <div class='sign-line'> </div>
    <div class='sign-line'>ลงชื่อ_______________________พยาน</div>
    <div class='sign-name'>(.........................................)</div>
</div>";
            }

            // ===== สร้างคอลัมน์ซ้ายจากลายเซ็นเดิม =====
            var leftColumnHtml = new StringBuilder();
            foreach (dynamic s in leftSigners)
                leftColumnHtml.Append(RenderSignatureBlock(s, isCompanySide: false));
            if (leftColumnHtml.Length == 0)
                leftColumnHtml.Append("<div class='sign-block placeholder'></div>");

            // ===== คอลัมน์ขวาเป็นแบบ fixed =====
            var rightFixedHtml = RenderFixedSignatureBlock();

            // ===== ประกอบตาราง 2 คอลัมน์ =====
            var signatory2ColHtml = $@"
<table class='signature-2col'>
  <tr>
    <td class='sign-col left' style='width:50%; vertical-align:top;'>
      {leftColumnHtml}
    </td>
    <td class='sign-col right' style='width:50%; vertical-align:top;'>
      {rightFixedHtml}
    </td>
  </tr>
</table>";



            // ── 5) เนื้อหา HTML (ใช้ ?? "" กัน null string) ────────────────────────
            string htmlBody = $@"
<div style='margin-bottom:24px;text-align:center;'>
    {(System.IO.File.Exists(logoPath) ? $"<img src='file:///{logoPath}' style='width:240px;height:80px;margin-bottom:24px;' />" : "")}
</div>
<div class='text-center t-16'><b>สัญญาจ้างลูกจ้าง</b></div>
<p class='tab3 t-14'>
    สัญญาฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ตำบล/แขวง ทุ่งสองห้อง อำเภอ/เขต หลักสี่ กรุงเทพมหานคร เมื่อ {strcontractsign}
</p>
<p class='tab3 t-14'>
    ระหว่าง สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม โดย {result.OSMEP_NAME ?? ""} ตำแหน่ง {result.OSMEP_POSITION ?? ""} {strAttorney} ซึ่งต่อไปในสัญญานี้จะเรียกว่า “ผู้ว่าจ้าง”
</p>
<p class='tab3 t-14'>
    ฝ่ายหนึ่ง กับ {result.EmploymentName ?? ""} เลขประจำตัวประชาชน {result.IdenID ?? ""} อยู่บ้านเลขที่ {result.EmpAddress ?? ""} ซึ่งต่อไปในสัญญานี้จะเรียกว่า “ลูกจ้าง” อีกฝ่ายหนึ่ง โดยทั้งสองฝ่ายได้ตกลงทำร่วมกันดังมีรายละเอียดต่อไปนี้
</p>
<p class='tab3 t-14'>
    1.ผู้ว่าจ้างตกลงจ้างลูกจ้างปฏิบัติงานกับผู้ว่าจ้าง โดยให้ปฏิบัติงานภายใต้งาน {result.WorkDetail ?? ""} ในตำแหน่ง {result.WorkPosition ?? ""} ปฏิบัติหน้าที่ ณ {result.Work_Location ?? ""} โดยมีรายละเอียดหน้าที่ความรับผิดชอบปรากฏตามเอกสารแนบท้ายสัญญาจ้าง ตั้งแต่ {strHiringStart} ถึง {strHiringEnd}
</p>
<p class='tab3 t-14'>
    2.ผู้ว่าจ้างจะจ่ายค่าจ้างให้แก่ลูกจ้างในระหว่างระยะเวลาการปฏิบัติงานของลูกจ้างตามสัญญานี้ในอัตราเดือนละ {result.Salary} บาท ({strSalary}) โดยจะจ่ายให้ในวันทำการก่อนวันทำการสุดท้ายของธนาคารในเดือนนั้นสามวันทำการ และนำเข้าบัญชีเงินฝากของลูกจ้าง ณ ที่ทำการของผู้ว่าจ้าง หรือ ณ ที่อื่นใดตามที่ผู้ว่าจ้างกำหนด
</p>
 <p class='tab3 t-14'>
    3.ในการจ่ายค่าจ้าง และ/หรือ เงินในลักษณะอื่นให้แก่ลูกจ้าง ลูกจ้างตกลงยินยอมให้ผู้ว่าจ้างหักภาษี ณ ที่จ่าย และ/หรือ เงินอื่นใดที่ต้องหักโดยชอบด้วยระเบียบ ข้อบังคับของผู้ว่าจ้างหรือตามกฎหมายที่เกี่ยวข้อง
</p>
<p class='tab3 t-14'>
    4.ตลอดระยะเวลาการปฏิบัติงานตามสัญญานี้ ลูกจ้างมีสิทธิได้รับสิทธิประโยชน์อื่น ๆ ตามที่กำหนดไว้ใน ระเบียบ ข้อบังคับ คำสั่ง หรือประกาศใด ๆ ตามที่ผู้ว่าจ้างกำหนด
</p>
<p class='tab3 t-14'>
    5.ผู้ว่าจ้างจะทำการประเมินผลการปฏิบัติงานอย่างน้อยปีละสองครั้ง ตามหลักเกณฑ์และวิธีการที่ผู้ว่าจ้างกำหนด ทั้งนี้ หากผลการประเมินไม่ผ่านตามหลักเกณฑ์ที่กำหนด ผู้ว่าจ้างมีสิทธิบอกเลิกสัญญาจ้างได้ และลูกจ้างไม่มีสิทธิเรียกร้องเงินชดเชยหรือเงินอื่นใด
</p>
<p class='tab3 t-14'>
    6.ตลอดระยะเวลาการปฏิบัติงานตามสัญญานี้ ลูกจ้างจะต้องปฏิบัติตามกฎ ระเบียบ ข้อบังคับ คำสั่งหรือประกาศใด ๆ ของผู้ว่าจ้าง ตลอดจนมีหน้าที่ต้องรักษาวินัยและยอมรับการลงโทษทางวินัยของผู้ว่าจ้างโดยเคร่งครัด และยินยอมให้ถือว่า กฎหมาย ระเบียบ ข้อบังคับ หรือคำสั่งต่าง ๆ ของผู้ว่าจ้างเป็นส่วนหนึ่งของสัญญาจ้างนี้
</p>
<p class='tab3 t-14'>
    ในกรณีลูกจ้างจงใจขัดคำสั่งโดยชอบของผู้ว่าจ้างหรือละเลยไม่นำพาต่อคำสั่งเช่นว่านั้นเป็นอาจิณ หรือประการอื่นใด อันไม่สมควรกับการปฏิบัติหน้าที่ของตนให้ลุล่วงไปโดยสุจริตและถูกต้อง ลูกจ้างยินยอมให้ผู้ว่าจ้างบอกเลิกสัญญาจ้างโดยมิต้องบอกกล่าวล่วงหน้า
</p>
<p class='tab3 t-14'>
    7. ลูกจ้างต้องปฏิบัติงานให้กับผู้ว่าจ้าง ตามที่ได้รับมอบหมายด้วยความซื่อสัตย์ สุจริต และตั้งใจปฏิบัติงานอย่างเต็มกำลังความสามารถของตน โดยแสวงหาความรู้และทักษะเพิ่มเติมหรือกระทำการใด 
    เพื่อให้ผลงานในหน้าที่มีคุณภาพดีขึ้น ทั้งนี้ ต้องรักษาผลประโยชน์และชื่อเสียงของผู้ว่าจ้าง และไม่เปิดเผยความลับหรือข้อมูลของทางราชการให้ผู้หนึ่งผู้ใดทราบ โดยมิได้รับอนุญาตจากผู้รับผิดชอบงานนั้น ๆ
</p>
<p class='tab3 t-14'>
    8. สัญญานี้สิ้นสุดลงเมื่อเข้ากรณีใดกรณีหนึ่ง ดังต่อไปนี้
</p>
<p class='tab4 t-14'>8.1 สิ้นสุดระยะเวลาตามสัญญาจ้าง</p>
<p class='tab4 t-14'>8.2 เมื่อผู้ว่าจ้างบอกเลิกสัญญาจ้าง หรือลูกจ้างบอกเลิกสัญญาจ้างตามข้อ 10</p>
<p class='tab4 t-14'>8.3 ลูกจ้างกระทำการผิดวินัยร้ายแรง</p>
<p class='tab4 t-14'>8.4 ลูกจ้างไม่ผ่านการประเมินผลการปฏิบัติงานของลูกจ้างตามข้อ 5</p>
<p class='tab3 t-14'>
    9. ในกรณีที่สัญญาสิ้นสุดตามข้อ 8.3 และ 8.4 ลูกจ้างยินยอมให้ผู้ว่าจ้างสั่งให้ลูกจ้างพ้นสภาพการเป็นลูกจ้างได้ทันที โดยไม่จำเป็นต้องมีหนังสือว่ากล่าวตักเตือน และผู้ว่าจ้างไม่ต้องจ่ายค่าชดเชยหรือเงินอื่นใดให้แก่ลูกจ้างทั้งสิ้น เว้นแต่ค่าจ้างที่ลูกจ้างจะพึงได้รับตามสิทธิ
</p>
<p class='tab3 t-14'>
    10. ลูกจ้างมีสิทธิบอกเลิกสัญญาจ้างได้ก่อนสัญญาครบกำหนด โดยทำหนังสือแจ้งเป็นลายลักษณ์อักษรต่อผู้ว่าจ้างได้ทราบล่วงหน้าไม่น้อยกว่า 30 วัน เมื่อผู้ว่าจ้างได้อนุมัติแล้ว ให้ถือว่าสัญญาจ้างนี้ได้สิ้นสุดลง
</p>
<p class='tab3 t-14'>
    11. ในกรณีที่ลูกจ้างกระทำการใดอันทำให้ผู้ว่าจ้างได้รับความเสียหาย ไม่ว่าเหตุนั้นผู้ว่าจ้างจะนำมาเป็นเหตุบอกเลิกสัญญาจ้างหรือไม่ก็ตาม ผู้ว่าจ้างมีสิทธิจะเรียกร้องค่าเสียหาย และลูกจ้างยินยอมชดใช้ค่าเสียหายตามที่ผู้ว่าจ้างเรียกร้องทุกประการ 
</p>
<p class='tab3 t-14'>
    12. ลูกจ้างจะต้องไม่เปิดเผยหรือบอกกล่าวอัตราค่าจ้างของลูกจ้างให้แก่บุคคลใดทราบ ไม่ว่าจะโดยวิธีใดหรือเวลาใด เว้นแต่จะเป็นการกระทำตามกฎหมายหรือคำสั่งศาล
</p>
<p class='tab3 t-14'>
    13. สัญญาฉบับนี้ได้จัดทำขึ้นเป็นสัญญาอิเล็กทรอนิกส์คู่สัญญาได้อ่านตรวจสอบและทำความเข้าใจ ข้อความในสัญญาฉบับนี้โดยละเอียดแล้วจึงได้ลงลายมือชื่ออิเล็กทรอนิกส์ไว้เป็นหลักฐาน ณ วัน เดือน ปี ดังกล่าวข้างต้น 
และมีพยานรู้ถึงการลงนามของคู่สัญญา และคู่สัญญาต่างฝ่ายต่างเก็บรักษาไฟล์สัญญาอิเล็กทรอนิกส์ฉบับนี้ไว้เป็นหลักฐาน
</p>
<!--<p class='tab3 t-14'>
    สัญญาฉบับนี้ได้จัดทำขึ้นเป็นสัญญาอิเล็กทรอนิกส์คู่สัญญาได้อ่านตรวจสอบและทำความเข้าใจข้อความในสัญญาฉบับนี้โดยละเอียดแล้ว จึงได้ลงลายมือชื่ออิเล็กทรอนิกส์ไว้เป็นหลักฐาน ณ วัน เดือน ปี ดังกล่าวข้างต้น 
และมีพยานรู้ถึงการลงนามของคู่สัญญา และคู่สัญญาต่างฝ่ายต่างเก็บรักษาไฟล์สัญญาอิเล็กทรอนิกส์ฉบับนี้ไว้เป็นหลักฐาน
</p> -->
<p class='tab3 t-14'>
     สัญญานี้ทำขึ้นเป็นสัญญาอิเล็กทรอนิกส์ คู่สัญญาได้อ่าน เข้าใจเงื่อนไข และยอมรับเงื่อนไข และได้ยืนยันว่าเป็นผู้มีอำนาจลงนามในสัญญาจึงได้ลงลายมืออิเล็กทรอนิกส์พร้อมทั้งประทับตรา (ถ้ามี) ในสัญญาไว้ และต่างฝ่ายต่างยึดถือไว้ฝ่ายละหนึ่งฉบับในระบบของตน 
</p>
</br>
</br>
{signatory2ColHtml} 



<div style='page-break-before: always;'></div>
<p class='text-center t-16' style='font-weight:bold;'>เอกสารแนบท้ายสัญญาจ้างลูกจ้าง</p>
<p class='text-center t-16' style='font-weight:bold;'>งานศูนย์ให้บริการ SMEs ครบวงจร</p>
<p class='tab2 t-14'>หน้าที่ความรับผิดชอบ : {result.Work_Detail ?? ""}</p>
<p class='tab2 t-14'>เจ้าหน้าที่ศูนย์ให้บริการ SMEs ครบวงจร กรุงเทพมหานคร</p>
<p class='tab2 t-14'>- การปรับปรุงข้อมูลผู้ประกอบการ SME (ไม่น้อยกว่า 30 ราย/เดือน)</p>
<p class='tab2 t-14'>- การให้บริการคำปรึกษา แนะนำทางธุรกิจ อาทิเช่น ด้านบัญชี การเงิน การตลาด การบริหารจัดการ การผลิต กฎหมาย เทคโนโลยีสารสนเทศ และอื่น ๆ ที่เกี่ยวข้องทางธุรกิจ (ไม่น้อยกว่า 30 ราย/เดือน)</p>
<p class='tab2 t-14'>- สนับสนุน เสนอแนะแนวทางการแก้ไขปัญหาให้ SME ได้รับประโยชน์ตามมาตรการของภาครัฐ</p>
<p class='tab2 t-14'>- สนับสนุนการพัฒนาเครือข่ายหน่วยงานให้บริการส่งเสริม SME ให้บริการส่งต่อภายใต้<br>หน่วยงานพันธมิตร การติดตามผลและประสานงานแก้ไขปัญหา</p>
<p class='tab2 t-14'>- สนับสนุนนโยบาย มาตรการ และการทำงานของ สสว. ในการสร้าง ประสาน เชื่อมโยงเครือข่ายในพื้นที่ (รูปแบบ Online & Offline) เพื่อสนับสนุนการปฏิบัติงานตามภารกิจ</p>
<p class='tab2 t-14'>- สนับสนุนจัดทำข้อมูล SME จังหวัด เพื่อนำข้อมูลมาใช้ประโยชน์ในการเสนอแนะทางธุรกิจแก่ SME และเชื่อมโยงไปสู่การแก้ปัญหาหรือการจัดทำมาตรการภาครัฐ</p>
<p class='tab2 t-14'>- ปฏิบัติงานภายใต้การบังคับบัญชาของผู้จัดการศูนย์กลุ่มจังหวัดฯ หรือ ผู้จัดการศูนย์ให้บริการ SMEs ครบวงจร กรุงเทพมหานคร ตามประกาศ สสว. และเข้าร่วมกิจกรรมต่าง ๆ </p>
<p class='tab2 t-14'>- กำกับดูแลข้อมูลตาม พ.ร.บ.การคุ้มครองข้อมูลส่วนบุคคล</p>
<p class='tab2 t-14'>- งานอื่น ๆ ตามที่ได้รับมอบหมาย</p>
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
         body {{
            font-size: 22px;
            font-family: 'TH Sarabun PSK', Arial, sans-serif;
         
        }}
        /* แก้การตัดคำไทย: ไม่หั่นกลางคำ, ตัดเมื่อจำเป็น */
        body, p, div {{
            word-break: keep-all;            /* ห้ามตัดกลางคำ */
            overflow-wrap: break-word;       /* ตัดเฉพาะเมื่อจำเป็น (ยาวจนล้นบรรทัด) */
            -webkit-line-break: after-white-space; /* ช่วย WebKit เก่าจัดบรรทัด */
            hyphens: none;
        }}
        /* ตารางลายเซ็นแบบ 2 คอลัมน์ กึ่งกลางหน้า */
        .signature-2col{{
            width: 90%;
            max-width: 800px;  /* A4 portrait ดูบาลานซ์ */
            margin: 24px auto; /* กึ่งกลางหน้า */
            table-layout: fixed;
            border-collapse: separate;
           font-size: 1.1em;
            font-family: 'TH Sarabun PSK', Arial, sans-serif;
        }}
        .signature-2col .sign-col{{
            width: 50%;
            vertical-align: top;
            padding: 8px 12px;
            text-align: center;         /* กึ่งกลางในคอลัมน์ */
        }}
        .sign-block{{
            margin-top: 12px;
            font-size: inherit;       /* ใช้ขนาดฟอนต์จาก .signature-2col */
            font-family: inherit;     /* ใช้ฟอนต์จาก .signature-2col */
            page-break-inside: avoid;   /* กันบล็อกถูกตัดครึ่งหน้า */
        }}
        .sign-block.placeholder{{
            height: 120px;
        }}

        .sign-img img{{
            max-height: 80px;
            display: inline-block;
        }}
        .seal-img img{{
            max-height: 80px;
            display: inline-block;
            margin-top: 6px;
        }}
        .seal-caption{{
            margin-top: 4px; 
            font-size: inherit;       /* ใช้ขนาดฟอนต์จาก .signature-2col */
            font-family: inherit;     /* ใช้ฟอนต์จาก .signature-2col */
        }}

        .sign-line{{
            font-size: inherit;       /* ใช้ขนาดฟอนต์จาก .signature-2col */
            font-family: inherit;     /* ใช้ฟอนต์จาก .signature-2col */
            text-align: center;
        }}
        .sign-name, .sign-unit{{
            margin-top: 4px;
            text-align: center;
            font-size: inherit;       /* ใช้ขนาดฟอนต์จาก .signature-2col */
            font-family: inherit;     /* ใช้ฟอนต์จาก .signature-2col */
        }}

       .t-12 {{ font-size: 1em; }}
        .t-14 {{ font-size: 1.1em; }}
        .t-16 {{ font-size: 1.5em; }}
        .t-18 {{ font-size: 1.7em; }}
        .t-20 {{ font-size: 1.9em; }}
        .t-22 {{ font-size: 2.1em; }}

        .tab1 {{ text-indent: 48px;}}
        .tab2 {{ text-indent: 96px;}}
        .tab3 {{ text-indent: 144px;}}
        .tab4 {{ text-indent: 192px;}}
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
            word-break: break-all;
        }}
        .text-right {{ text-align: right; }}
        .contract, . {{
            margin: 12px 0;
            line-height: 1.7;
        }}
        . {{
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
