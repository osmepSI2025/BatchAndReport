using BatchAndReport.DAO;
using BatchAndReport.Entities;
using BatchAndReport.Models;
using BatchAndReport.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;


public class WordEContract_HireEmployee
{
    private readonly WordServiceSetting _w;
    private readonly Econtract_Report_ECDAO _e;
    public WordEContract_HireEmployee(WordServiceSetting ws, Econtract_Report_ECDAO e)
    {
        _w = ws;
        _e = e;
    }
    #region   4.1.3.3. สัญญาจ้างลูกจ้าง
    public async Task<byte[]> OnGetWordContact_HireEmployee(string id)
    {
        try {
            var result =await _e.GetECAsync(id);
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
                string strcontractsign = CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now);
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สัญญาฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เลขที่ 21 ถนนวิภาวดีรังสิต เขตจตุจักร กรุงเทพมหานคร เมื่อ"+ strcontractsign + "", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ระหว่าง สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม โดย........................................." +
                    "ผู้อำนวยการฝ่ายศูนย์ให้บริการ SMEs ครบวงจร สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ผู้รับมอบหมายตามคำสั่งสำนักงานฯ ที่ 629/2564 ลงวันที่ 30 กันยายน 2564 ซึ่งต่อไปในสัญญานี้จะเรียกว่า “ผู้ว่าจ้าง”", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ฝ่ายหนึ่ง กับ "+result.SignatoryName+" เลขประจำตัวประชาชน " + result.IdenID + " อยู่บ้านเลขที่ "+result.EmpAddress+" " +
                    "ซึ่งต่อไปในสัญญานี้จะเรียกว่า “ลูกจ้าง” อีกฝ่ายหนึ่ง โดยทั้งสองฝ่ายได้ตกลงทำร่วมกันดังมีรายละเอียดต่อไปนี้", null, "32"));

                string strHiringStart = CommonDAO.ToThaiDateStringCovert(result.HiringStartDate??DateTime.Now);
                string strHiringEnd = CommonDAO.ToThaiDateStringCovert(result.HiringEndDate ?? DateTime.Now);
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("1.ผู้ว่าจ้างตกลงจ้างลูกจ้างปฏิบัติงานกับผู้ว่าจ้าง โดยให้ปฏิบัติงานภายใต้งาน "+result.WorkDetail+"  ในตำแหน่ง "+result.WorkPosition+" ปฏิบัติหน้าที่ ณ ศูนย์กลุ่มจังหวัดให้บริการ SME ครบวงจร ..... " +
                    "โดยมีรายละเอียดหน้าที่ความรับผิดชอบปรากฏตามเอกสารแนบท้ายสัญญาจ้าง ตั้งแต่"+ strHiringStart + " ถึง"+ strHiringEnd + "", null, "32"));

                string strSalary = CommonDAO.NumberToThaiText(result.Salary??0);
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("2.ผู้ว่าจ้างจะจ่ายค่าจ้างให้แก่ลูกจ้างในระหว่างระยะเวลาการปฏิบัติงานของลูกจ้างตามสัญญานี้ในอัตราเดือนละ "+result.Salary+"บาท ("+ strSalary + ")" +
                    "โดยจะจ่ายให้ในวันทำการก่อนวันทำการสุดท้ายของธนาคารในเดือนนั้นสามวันทำการ และนำเข้าบัญชีเงินฝากของลูกจ้าง ณ ที่ทำการของผู้ว่าจ้าง หรือ ณ ที่อื่นใดตามที่ผู้ว่าจ้างกำหนด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.ในการจ่ายค่าจ้าง และ/หรือ เงินในลักษณะอื่นให้แก่ลูกจ้าง ลูกจ้างตกลงยินยอมให้ผู้ว่าจ้างหักภาษี ณ ที่จ่าย และ/หรือ เงินอื่นใดที่ต้องหักโดยชอบด้วยระเบียบ ข้อบังคับของผู้ว่าจ้างหรือตามกฎหมายที่เกี่ยวข้อง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("4.ตลอดระยะเวลาการปฏิบัติงานตามสัญญานี้ ลูกจ้างมีสิทธิได้รับสิทธิประโยชน์อื่น ๆ ตามที่กำหนดไว้ใน ระเบียบ ข้อบังคับ คำสั่ง หรือประกาศใด ๆ ตามที่ผู้ว่าจ้างกำหนด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("5.ผู้ว่าจ้างจะทำการประเมินผลการปฏิบัติงานอย่างน้อยปีละสองครั้ง ตามหลักเกณฑ์และวิธีการที่ผู้ว่าจ้างกำหนด ทั้งนี้ หากผลการประเมินไม่ผ่านตามหลักเกณฑ์ที่กำหนด ผู้ว่าจ้างมีสิทธิบอกเลิกสัญญาจ้างได้ และลูกจ้างไม่มีสิทธิเรียกร้องเงินชดเชยหรือเงินอื่นใด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("6.ตลอดระยะเวลาการปฏิบัติงานตามสัญญานี้ ลูกจ้างจะต้องปฏิบัติตามกฎ ระเบียบ ข้อบังคับ คำสั่งหรือประกาศใด ๆ ของผู้ว่าจ้าง " +
                    "ตลอดจนมีหน้าที่ต้องรักษาวินัยและยอมรับการลงโทษทางวินัยของผู้ว่าจ้างโดยเคร่งครัด และยินยอมให้ถือว่า กฎหมาย ระเบียบ ข้อบังคับ หรือคำสั่งต่าง ๆ ของผู้ว่าจ้างเป็นส่วนหนึ่งของสัญญาจ้างนี้", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ในกรณีลูกจ้างจงใจขัดคำสั่งโดยชอบของผู้ว่าจ้างหรือละเลยไม่นำพาต่อคำสั่งเช่นว่านั้นเป็นอาจิณ หรือประการอื่นใด อันไม่สมควรกับการปฏิบัติหน้าที่ของตนให้ลุล่วงไปโดยสุจริตและถูกต้อง ลูกจ้างยินยอมให้ผู้ว่าจ้างบอกเลิกสัญญาจ้างโดยมิต้องบอกกล่าวล่วงหน้า", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("7. ลูกจ้างต้องปฏิบัติงานให้กับผู้ว่าจ้าง ตามที่ได้รับมอบหมายด้วยความซื่อสัตย์ สุจริต และตั้งใจปฏิบัติงานอย่างเต็มกำลังความสามารถของตน โดยแสวงหาความรู้และทักษะเพิ่มเติมหรือกระทำการใด " +
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
                body.AppendChild(WordServiceSetting.NormalParagraphWith_1Tabs("- การให้บริการคำปรึกษา แนะนำทางธุรกิจ อาทิเช่น ด้านบัญชี การเงิน การตลาด การบริหารจัดการ การผลิต กฎหมาย เทคโนโลยีสารสนเทศ และอื่น ๆ ที่เกี่ยวข้องทางธุรกิจ (ไม่น้อยกว่า 30 ราย/เดือน)", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_1Tabs("- สนับสนุน เสนอแนะแนวทางการแก้ไขปัญหาให้ SME ได้รับประโยชน์ตามมาตรการของภาครัฐ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_1Tabs("- สนับสนุนการพัฒนาเครือข่ายหน่วยงานให้บริการส่งเสริม SME ให้บริการส่งต่อภายใต้หน่วยงานพันธมิตร การติดตามผลและประสานงานแก้ไขปัญหา", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_1Tabs("- สนับสนุนนโยบาย มาตรการ และการทำงานของ สสว. ในการสร้าง ประสาน เชื่อมโยงเครือข่ายในพื้นที่ (รูปแบบ Online & Offline) เพื่อสนับสนุนการปฏิบัติงานตามภารกิจ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_1Tabs("- สนับสนุนจัดทำข้อมูล SME จังหวัด เพื่อนำข้อมูลมาใช้ประโยชน์ในการเสนอแนะทางธุรกิจแก่ SME และเชื่อมโยงไปสู่การแก้ปัญหาหรือการจัดทำมาตรการภาครัฐ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_1Tabs("- ปฏิบัติงานภายใต้การบังคับบัญชาของผู้จัดการศูนย์กลุ่มจังหวัดฯ หรือ ผู้จัดการศูนย์ให้บริการ SMEs ครบวงจร กรุงเทพมหานคร ตามประกาศ สสว. และเข้าร่วมกิจกรรมต่าง ๆ ", null, "32"));
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
                    new PageSize() { Width = 11906, Height = 16838 }, // A4 size
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
    #endregion    4.1.3.3. สัญญาจ้างลูกจ้าง

}
