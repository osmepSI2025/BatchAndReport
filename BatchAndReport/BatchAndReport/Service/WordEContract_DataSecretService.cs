using BatchAndReport.DAO;
using BatchAndReport.Models;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

public class WordEContract_DataSecretService
{
    private readonly WordServiceSetting _w;
    E_ContractReportDAO _eContractReportDAO;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    public WordEContract_DataSecretService(WordServiceSetting ws
         , E_ContractReportDAO eContractReportDAO
          , IConverter pdfConverter
        )
    {
        _w = ws;
        _eContractReportDAO = eContractReportDAO;
         _pdfConverter = pdfConverter;
    }
    #region  4.1.1.2.7.สัญญาการรักษาข้อมูลที่เป็นความลับ
    public async Task<byte[]> OnGetWordContact_DataSecretService(string id)
    {
        var result = await _eContractReportDAO.GetNDAAsync(id);
        if (result == null)
        {
            throw new Exception("NDA data not found.");
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
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("สัญญาการรักษาข้อมูลที่เป็นความลับ", "36"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("(Non-disclosure Agreement : NDA)", "36"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("ระหว่าง", "36"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม", "36"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("กับ" + result.Contract_Party_Name, "36"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("---------------------------------------------", "36"));

                // --- 3. Main contract body ---
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs(
                  "สัญญาการรักษาข้อมูลที่เป็นความลับ (“สัญญา”) ฉบับนี้จัดขึ้น เมื่อวันที่" +
                  "" + result.Sign_Date + "" +
                  "ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.) ระหว่าง", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.) โดย " + result.OSMEP_Signatory + "" +
                  "ตำแหน่ง " + result.OSMEP_Position + " เลขที่ 120 หมู่ 3 ศูนย์ราชการเฉลิมพระเกียรติ 80 พรรษา 5 ธันวาคม 2550. (อาคารซี) ชั้น 2, 10, 11 ถนนแจ้งวัฒนะ แขวงทุ่งสองห้อง เขตหลักสี่ กรุงเทพ 10210 ซึ่งต่อไปในสัญญานี้จะเรียกว่า “ผู้เปิดเผยข้อมูล”" +
                  "กับ กับ" + result.Contract_Party_Name + " โดย " + result.CP_Signatory + "" +
                  "ตำแหน่ง " + result.CP_Position + " สำนักงานตั้งอยู่เลขที่ " + result.OfficeLoc + " ซึ่งต่อไปในสัญญานี้จะเรียกว่า “ผู้รับข้อมูล”",
                  null, "32"
                ));


                // --- 5. NDA Clauses (example, add more as needed) ---
                body.AppendChild(WordServiceSetting.JustifiedParagraph(
                  "คู่สัญญาได้ตกลงทำสัญญากันมีข้อความดังต่อไปนี้", "32"
                ));

                var conPurpose = await _eContractReportDAO.GetNDA_RequestPurposeAsync(id);
                if (conPurpose != null && conPurpose.Count > 0)
                {
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 1 วัตถุประสงค์", null, "32", true));
                    body.AppendChild(WordServiceSetting.JustifiedParagraph_2tab(
                 "โดยที่ผู้ให้ข้อมูลเป้นเจ้าของข้อมูล ผู้รับข้อมูลมีความต้องการที่จะใช้ข้อมูลของผู้ให้ข้อมูลเพื่อที่จะดําเนินการตามวัตถุประสงค์ ดังนี้ (ระบุวัตถุประสงค์ของการขอข้อมูลที่เป็นความลับไปเพื่อใช้งาน)", "32"
               ));
                    foreach (var purpose in conPurpose)
                    {
                        body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs(purpose.Detail, null, "32"));
                    }
                }

                body.AppendChild(WordServiceSetting.JustifiedParagraph_2tab(
                  "โดยผู้รับข้อมูลจะไม่เปิดเผยข้อมูลดังกล่าวแก่บุคคลภายนอก เว้นแต่ได้รับความยินยอมเป็นลายลักษณ์อักษรจากผู้เปิดเผยข้อมูลก่อน ทั้งนี้ (ระบุเงื่อนไขหรือข้อยกเว้นตามข้อตกลง)", "32"
                ));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 2 ข้อมูลที่เป็นความลับ", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs(
               "“ข้อมูลที่เป็นความลับ” หมายความว่า บรรดาข้อความเอกสารข้อมูลตลอดจนรายละเอียดทั้งปวงที่เป็นของผู้ให้ข้อมูล" +
               "รวมถึงที่อยู่ในความครอบครองหรือควบคุมดูแลของผู้ให้ข้อมูล และไม่เป็นที่รับรู้ของสาธารณชนโดยทั่วไปไม่ว่าจะในรูปแบบที่จับต้องได้หรือไม่ก็ตาม" +
               "หรือสื่อแบบใดไม่ว่าจะถูกดัดแปลงแก้ไขโดยผู้รับข้อมูลหรือไม่ และไม่ว่าจะเปิดเผยเมื่อใดและอย่างไร ให้ถือว่าเป็นความลับโดยข้อมูลที่เป็นความลับอาจหมายความรวมถึง" +
               "ข้อมูลเชิงกลยุทธ์ของผู้ให้ข้อมูล แผนธุรกิจ ข้อมูลทางการเงิน ข้อมูลลูกจ้าง ข้อมูลผู้ประกอบการ เเละข้อมูลส่วนบุคคลที่ผู้ให้ข้อมูลได้เก็บ รวบรวม ใช้" +
               "ข้อมูลที่เป็นความลับที่ผู้ให้ข้อมูล หรือในนามของผู้ให้ข้อมูลที่เปิดเผยแก่ผู้รับข้อมูล ซึ่งหมายความรวมถึงข้อมูลที่ผู้ให้ข้อมูลให้แก่ผู้รับข้อมูล ดังนี้",
                 null,
               "32"
             ));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(ระบุประเภทของข้อมูลที่เป็นความลับที่นำส่งให้แก่กัน)", JustificationValues.Left, "32"));
                var conConfidentialType = await _eContractReportDAO.GetNDA_ConfidentialTypeAsync(id);
                if (conConfidentialType != null && conConfidentialType.Count > 0)
                {
                    foreach (var confidential in conConfidentialType)
                    {
                        body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs(confidential.Detail, null, "32"));
                    }
                }

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 3 การรักษาข้อมูลที่เป็นความลับ", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.1 ผู้รับข้อมูลต้องรับผิดชอบรักษาข้อมูลที่เป็นความลับและเก็บข้อมูลความลับไว้โดยครบถ้วนและอย่างเคร่งครัด ผู้รับข้อมูลจะต้องไม่เปิดเผยทําสําเนาหรือทําการอื่นใดทํานองเดียวกันแก่บุคคลอื่นไม่ว่าทั้งหมดหรือบางส่วน เว้นแต่ได้รับอนุญาตเป็นหนังสือจากผู้ให้ข้อมูล", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.2 ผู้รับข้อมูลต้องใช้ข้อมูลที่เป็นความลับเพื่อการอันเกี่ยวกับหรือสัมพันธ์กับการดําเนินงานที่มีอยู่ระหว่างผู้ให้ข้อมูลกับผู้รับข้อมูล โดยผู้รับข้อมูลต้องแจ้งให้ผู้ให้ข้อมูลทราบโดยทันทีที่พบการใช้หรือการเปิดเผยข้อมูลที่เป็นความลับโดยไม่ได้รับอนุญาตหรือการละเมิดหรือฝ่าฝืนข้อกําหนดตามสัญญานี้ อีกทั้ง ผู้รับข้อมูลจะต้องให้ความร่วมมือกับผู้ให้ข้อมูลอย่างเต็มที่ในการเรียกคืนซึ่งการครอบครองข้อมูลที่เป็นความลับ การป้องกันการใช้ข้อมูลที่เป็นความลับโดยไม่ได้รับอนุญาตและการระงับยับยั้งการเผยแพร่ข้อมูลที่เป็นความลับออกสู่สาธารณะ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.3 ผู้รับข้อมูลต้องจัดให้มีและคงไว้ซึ่งมาตรการรักษาความปลอดภัยสำหรับการจัดเก็บและประมวลผลข้อมูลที่มีความเหมาะสมในมาตรการเชิงองค์กร มาตรการเชิง" +
                    "เทคนิค และมาตรการเชิงกายภาพ โดยคำนึงถึงลักษณะ ขอบเขต และวัตถุประสงค์ของการดำเนินการตามวัตถุประสงค์ที่ของสัญญาฉบับนี้เป็นสำคัญ " +
                    "เพื่อป้องกันมิให้ข้อมูลที่เป็นความลับถูกนําไปใช้โดยมิได้รับอนุญาตหรือถูกเปิดเผยแก่บุคคลอื่น " +
                    "โดยผู้รับข้อมูลต้องใช้มาตรการการเก็บรักษาข้อมูลที่เป็นความลับในระดับเดียวกันกับที่ผู้รับข้อมูลใช้กับข้อมูลที่เป็นความลับของตนเองซึ่งต้องไม่น้อยกว่าการดูแลที่สมควร", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.4 ผู้รับข้อมูลต้องแจ้งให้บุคลากร พนักงาน ลูกจ้าง ที่ปรึกษาของผู้รับข้อมูลและ/หรือบุคคลภายนอกที่ต้องเกี่ยวข้องกับข้อมูลที่เป็นความลับนั้น " +
                    "ทราบถึงความเป็นความลับและข้อจํากัดสิทธิในการใช้และการเปิดเผยข้อมูลที่เป็นความลับ และผู้รับข้อมูลต้องดําเนินการให้บุคคลดังกล่าวต้องผูกพันด้วยสัญญาหรือข้อตกลงเป็นหนังสือในการรักษาข้อมูลที่เป็นความลับ " +
                    "โดยมีข้อกําหนดเช่นเดียวกับหรือไม่น้อยกว่าข้อกําหนดและเงื่อนไขในสัญญาฉบับนี้ด้วย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.5 ข้อมูลที่เป็นความลับตามสัญญาฉบับนี้ไม่รวมไปถึงข้อมูลดังต่อไปนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(1) ข้อมูลที่ผู้ให้ข้อมูลเปิดเผยแก่สาธารณะ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(2) ข้อมูลที่ผู้รับข้อมูลทราบอยู่ก่อนที่ผู้ให้ข้อมูลจะเปิดเผยข้อมูลนั้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(3) ข้อมูลที่มาจากการพัฒนาโดยอิสระของผู้รับข้อมูลเอง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(4) ข้อมูลที่ต้องเปิดเผยโดยกฎหมายหรือตามคําสั่งศาล ทั้งนี้ ผู้รับข้อมูลต้องมีหนังสือแจ้งผู้ให้ข้อมูลได้รับทราบถึงข้อกําหนดหรือคําสั่งดังกล่าว โดยแสดงเอกสารข้อกำหนด หมายศาลและ/หรือหมายค้นอย่างเป็นทางการต่อผู้ให้ข้อมูลก่อนที่จะดําเนินการเปิดเผยข้อมูลดังกล่าว และในการเปิดเผยข้อมูลดังกล่าว ผู้รับข้อมูลจะต้องดําเนินการตามขั้นตอนทางกฎหมายเพื่อขอให้คุ้มครองข้อมูลดังกล่าวไม่ให้ถูกเปิดเผยต่อสาธารณะด้วย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(5) ผู้รับข้อมูลได้รับความยินยอมเป็นลายลักษณ์อักษรให้เปิดเผยข้อมูลจากผู้ให้ข้อมูล ก่อนที่ผู้รับข้อมูลจะเปิดเผยข้อมูลนั้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(6) ผู้รับข้อมูลได้รับข้อมูลที่เป็นความลับจากบุคคลที่สามที่ไม่อยู่ภายใต้ข้อกำหนดในเรื่องการรักษาความลับ หรือข้อจำกัดในเรื่องสิทธิ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.6 ผู้รับข้อมูลต้องไม่ทำซ้ำข้อมูลที่เป็นความลับแม้เพียงส่วนหนึ่งส่วนใดหรือทั้งหมด เว้นแต่การทำซ้ำเพื่อการใช้ข้อมูลที่เป็นความลับให้บรรลุผลตามวัตถุประสงค์ที่กำหนดไว้ในสัญญานี้ และไม่ทำวิศวกรรมย้อนกลับ หรือถอดรหัสข้อมูลที่เป็นความลับ ต้นแบบ หรือสิ่งอื่นใดที่บรรจุข้อมูลที่เป็นความลับ รวมทั้งไม่เคลื่อนย้าย พิมพ์ทับ หรือทำให้เสียรูปซึ่งสัญลักษณ์ที่แสดงเครื่องหมายสิทธิบัตร อนุสิทธิบัตร ลิขสิทธิ์ เครื่องหมายการค้า ตราสัญลักษณ์ และเครื่องหมายอื่นใดที่แสดงกรรมสิทธิ์ของต้นแบบหรือสำเนาของข้อมูลที่เป็นความลับที่ได้รับมาจากผู้ให้ข้อมูล", null, "32"));



                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 4 ทรัพย์สินทางปัญญา", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("สัญญาฉบับนี้ไม่มีผลบังคับใช้เป็นการโอนสิทธิหรือการอนุญาตให้ใช้สิทธิ (ไม่ว่าโดยตรง หรือโดยอ้อม) ให้แก่ผู้รับข้อมูลที่ได้รับความลับซึ่งสิทธิบัตร ลิขสิทธิ์ การออกแบบ เครื่องหมายการค้า ตราสัญลักษณ์ รูปประดิษฐ์อื่นใด ชื่อทางการค้า ความลับทางการค้า ไม่ว่าจดทะเบียนไว้ตามกฎหมายหรือไม่ก็ตาม หรือ สิทธิอื่น ๆ ของผู้ให้ข้อมูล ซึ่งอาจปรากฏอยู่หรือนํามาทําซ้ำไว้ในเอกสารข้อมูลที่เป็นความลับ ทั้งนี้ ผู้รับข้อมูลหรือบุคคลอื่นใดที่เกี่ยวข้องกับผู้รับข้อมูลและเกี่ยวข้องกับข้อมูลที่เป็นความลับดังกล่าวจะไม่ยื่นขอรับสิทธิและ/หรือขอจดทะเบียนเกี่ยวกับทรัพย์สินทางปัญญาใด ๆ ตลอดจนไม่นําไปใช้โดยไม่ได้รับการอนุญาตเป็นลายลักษณ์อักษรจากผู้ให้ข้อมูลเกี่ยวกับรายละเอียดข้อมูลที่เป็นความลับหรือส่วนหนึ่งส่วนใดของรายละเอียดดังกล่าว", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 5 การส่งคืน ลบ หรือการทําลายข้อมูลที่เป็นความลับ", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เมื่อการดําเนินงานที่มีอยู่ระหว่างผู้ให้ข้อมูลกับผู้รับข้อมูลเสร็จสิ้นลงตามวัตถุประสงค์ผู้รับข้อมูลจะต้องส่งมอบข้อมูลที่เป็นความลับและสําเนาของข้อมูลที่เป็นความลับที่ผู้รับข้อมูลได้รับไว้คืนให้แก่ผู้ให้ข้อมูล เว้นแต่ผู้ให้ข้อมูลเห็นว่าไม่ต้องนำส่งคืนแต่ต้องเลิกใช้ข้อมูลที่เป็นความลับ และทำการลบหรือทําลายข้อมูลที่เป็นความลับทั้งถูกจัดเก็บไว้ในคอมพิวเตอร์หรืออุปกรณ์อื่นใดที่ใช้จัดเก็บข้อมูล (ถ้ามี) หรือดําเนินการอื่นตามที่ได้รับการแจ้งเป็นลายลักษณ์อักษรจากผู้ให้ข้อมูล", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 6 การชดใช้ค่าเสียหาย", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("6.1 กรณีที่ผู้รับข้อมูล พนักงาน ลูกจ้าง ที่ปรึกษาของผู้รับข้อมูล และ/หรือบุคคลภายนอกที่ได้รับข้อมูลที่เป็นความลับจากผู้รับข้อมูลฝ่าฝืนข้อกำหนดตามสัญญานี้และก่อให้เกิดความเสียหายแก่ผู้ให้ข้อมูล และ/หรือบุคคลอื่น ผู้รับข้อมูลจะต้องชดใช้ค่าเสียหายให้แก่ผู้ให้ข้อมูล และ/หรือบุคคลที่ได้รับความเสียหายสำหรับความเสียหายเช่นว่านั้น ทั้งนี้ ผู้รับข้อมูลจะต้องเเจ้งเเก่ผู้ให้ข้อมูลทราบเป็นลายลักษณ์อักษรภายใน 7 วันนับตั้งเเต่มีการละเมิดข้อมูลที่เป็นความลับเกิดขึ้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("6.2 ผู้รับข้อมูลรับทราบว่าการเปิดเผยหรือการใช้ข้อมูลที่เป็นความลับโดยฝ่าฝืนข้อกำหนดตามสัญญานี้จะก่อให้เกิดความเสียหายแก่ผู้ให้ข้อมูลในจำนวนที่ไม่สามารถประเมินได้ ดังนั้น ผู้รับข้อมูลยินยอมให้ผู้ให้ข้อมูลใช้สิทธิที่จะร้องขอต่อศาลเพื่อให้มีคำสั่งให้ผู้รับข้อมูลหยุดการกระทำใด ๆ ที่เป็นการฝ่าฝืนข้อกำหนดตามสัญญานี้ และ/หรือใช้วิธีคุ้มครองชั่วคราวใด ๆ ตามที่ผู้ให้ข้อมูลเห็นว่าเหมาะสมได้ โดยผู้รับข้อมูลจะเป็นผู้รับผิดชอบค่าใช้จ่ายต่าง ๆ ที่เกิดขึ้นทั้งหมดจากการดำเนินการดังกล่าว", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("6.3 กรณีที่ผู้ให้ข้อมูลสงสัยว่าผู้รับข้อมูลฝ่าฝืนข้อกำหนดตามสัญญานี้ ผู้รับข้อมูลจะต้องเป็นฝ่ายพิสูจน์ว่าผู้รับข้อมูลไม่ได้ฝ่าฝืนข้อกำหนดตามสัญญานี้", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 7 ระยะเวลาตามสัญญา", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("สัญญานี้มีผลบังคับใช้นับตั้งแต่วันที่ทำสัญญานี้ โดยมีกำหนดระยะเวลาทั้งสิ้น " + result.EnforcePeriods + "ปี นับตั้งแต่วันที่ทำสัญญาฉบับนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เมื่อครบกำหนดระยะเวลาตามวรรคหนึ่ง หรือเมื่อมีการบอกเลิกสัญญา หรือผู้ให้ข้อมูลได้แจ้งให้ผู้รับข้อมูลดำเนินการทำลายข้อมูลดังกล่าว ผู้รับข้อมูลจะต้องดำเนินการทำลายข้อมูล ภายใน 7 วันนับเเต่ได้รับหนังสือร้องขอจากผู้ให้ข้อมูล ทั้งนี้ ผู้รับข้อมูลจะต้องไม่มีการสงวนไว้ซึ่งสำเนาใด ๆ", null, "32"));



                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 8  ข้อตกลงอื่น ๆ", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("8.1 ในกรณีที่มีเหตุจำเป็นต้องมีการเปลี่ยนแปลงแก้ไขสัญญานี้ ให้ทำเป็นลายลักษณ์อักษร และลงนามโดยคู่สัญญาหรือผู้มีอำนาจลงนามผูกพันนิติบุคคลและประทับตราสำคัญของนิติบุคคล (ถ้ามี) ของคู่สัญญา แล้วแต่กรณี", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("8.2 กรณีที่ผู้รับข้อมูลได้โอนกิจการ รวมกิจการ หรือควบกิจการ หรือดำเนินการอื่น ๆ ในลักษณะที่มีการเปลี่ยนแปลงของวัตถุประสงค์ในการดำเนินกิจการของผู้รับข้อมูลผู้รับข้อมูลจะต้องแจ้งให้ผู้ให้ข้อมูลทราบภายใน 5 วันทำการ นับแต่ได้เกิดเหตุดังกล่าวขึ้น", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 9 การบังคับใช้", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("9.1 ในกรณีที่ปรากฏภายหลังว่าส่วนใดส่วนหนึ่งในสัญญาฉบับนี้เป็นโมฆะให้ถือว่าข้อกําหนดส่วนที่เป็นโมฆะไม่มีผลบังคับในสัญญานี้ และข้อกําหนดที่เหลืออยู่ในสัญญาฉบับนี้ยังคงใช้บังคับและมีผลอยู่อย่างสมบูรณ์", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("9.2 สัญญาฉบับนี้อยู่ภายใต้การบังคับและตีความตามกฎหมายของประเทศไทย ให้ศาลของประเทศไทยมีอำนาจในกรณีที่มีข้อพิพาทใด ๆ อันเกิดขึ้นจากสัญญาฉบับนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("สัญญานี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน คู่สัญญาได้อ่าน และเข้าใจข้อความในสัญญาโดยละเอียดตลอดแล้ว เห็นว่าตรงตามเจตนารมณ์ทุกประการ จึงได้ลงลายมือชื่อพร้อมทั้งประทับตราสำคัญผูกพันนิติบุคคล (ถ้ามี) ไว้เป็นสำคัญ ณ วัน เดือน ปี ที่ระบุข้างต้น และคู่สัญญาต่างฝ่ายต่างยึดถือไว้ฝ่ายละหนึ่งฉบับ", null, "32"));

                // --- 6. Signature lines ---

                body.AppendChild(WordServiceSetting.EmptyParagraph());

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
                    // Row 1: Main signatures
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.NormalParagraph("ลงชื่อ............................................................ผู้ให้ข้อมูล", JustificationValues.Left, "32"),
                            WordServiceSetting.CenteredParagraph("(สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม)", "32")
                        ),
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.NormalParagraph("ลงชื่อ............................................................ผู้รับข้อมูล", JustificationValues.Left, "32"),
                            WordServiceSetting.CenteredBoldColoredParagraph("(" + result.Contract_Party_Name + ")", "000000", "32")
                        )
                    ),
                    // Row 2: Witnesses
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                                                   WordServiceSetting.NormalParagraph("ลงชื่อ............................................................พยาน", JustificationValues.Left, "32"),
                            WordServiceSetting.CenteredParagraph("(สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม)", "32")
                        ),
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.NormalParagraph("ลงชื่อ............................................................พยาน", JustificationValues.Left, "32"),
                              WordServiceSetting.CenteredBoldColoredParagraph("(" + result.Contract_Party_Name + ")", "000000", "32")
                        )
                    )
                );

                // Add signature table
                body.AppendChild(signatureTable);



                // --- 7. Add header/footer if needed ---
                WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);
            }
            stream.Position = 0;
            return stream.ToArray();
        }

    }

    public async Task<string> OnGetWordContact_DataSecretService_ToPDF(string id,string typeContact) 
    {
        var result = await _eContractReportDAO.GetNDAAsync(id);
        var conPurpose = await _eContractReportDAO.GetNDA_RequestPurposeAsync(id);
        
        var purposeHtml = "";
        if (conPurpose != null && conPurpose.Count > 0)
        {
            purposeHtml += "<p class='tab3 t-16'><b>ข้อ ๑ วัตถุประสงค์</b></p>";
            purposeHtml += "<p class='tab3 t-16'>โดยที่ผู้ให้ข้อมูลเป็นเจ้าของข้อมูล ผู้รับข้อมูลมีความต้องการที่จะใช้ข้อมูลของผู้</br>ให้ข้อมูลเพื่อที่จะดำเนินการตามวัตถุประสงค์ ดังนี้ </p>";
            foreach (var purpose in conPurpose)
            {
                purposeHtml += $"<p class='tab4 t-16'>{System.Net.WebUtility.HtmlEncode(purpose.Detail)}</p>";
            }
        }
        var conConfidentialHtml = "";
        var conConfidentialType = await _eContractReportDAO.GetNDA_ConfidentialTypeAsync(id);
        if (conConfidentialType != null && conConfidentialType.Count > 0)
        {
            foreach (var confidential in conConfidentialType)
            {
                conConfidentialHtml += $"<div class='tab3 t-16'>{System.Net.WebUtility.HtmlEncode(confidential.Detail)}</div>";
            }
        }
        if (result == null)
        {
            throw new Exception("NDA data not found.");
        }
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");
        var cssPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "css", "contract.css").Replace("\\", "/");
        var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
        string logoBase64 = "";
        if (System.IO.File.Exists(logoPath))
        {
            var bytes = System.IO.File.ReadAllBytes(logoPath);
            logoBase64 = Convert.ToBase64String(bytes);
        }
        #region checkมอบอำนาจ
        string strAttorneyLetterDate = CommonDAO.ToArabicDateStringCovert(result.Grant_Date ?? DateTime.Now);
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
        if (result.CP_S_AttorneyFlag == true)
        {
            strAttorney = "ผู้มีอำนาจ กระทำการแทน ปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ ฉบับลงวันที่ " + strAttorneyLetterDate_CP + "";

        }
        else
        {
            strAttorney = "";
        }
        #endregion

        var strDateTH = CommonDAO.ToThaiDateStringCovert(result.Sign_Date ?? DateTime.Now);


        #region signlist 

        var signlist = await _eContractReportDAO.GetSignNameAsync(id, typeContact);
        var signatoryTableHtml = "";
        if (signlist.Count > 0)
        {
            signatoryTableHtml = await _eContractReportDAO.RenderSignatory(signlist);

        }
        #endregion signlist


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
</head><body>
    <!-- Top Row: Logo left, Contract code box right -->
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
            <div style='display:inline-block; padding:20px; font-size:32pt;'>
             <img src='data:image/jpeg;base64,{logoBase64}' width='240' height='80' />
            </div>
        </td>
    </tr>
</table>
</br>
    <!-- Titles -->
    <div class='text-center t-22'><B>สัญญาการรักษาข้อมูลที่เป็นความลับ</B></div>
    <div class='text-center t-22'><B>(Non-disclosure Agreement : NDA)</B></div>
    <div class='text-center t-22'><B>ระหว่าง</B></div>
    <div class='text-center t-22'><B>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B></div>
    <div class='text-center t-22'><B>กับ {result.Contract_Party_Name}</B></div>
    <div class='text-center  t-22'>---------------------------------------------</div>
</br>
    <!-- Main contract body -->
   <p class='tab3 t-16'>
        สัญญาการรักษาข้อมูลที่เป็นความลับ (“สัญญา”) ฉบับนี้จัดขึ้น เมื่อ {strDateTH} ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.) ระหว่าง
    </p>
   <p class='tab3 t-16'>
        สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B>  โดย {result.OSMEP_NAME} ตำแหน่ง {result.OSMEP_POSITION} {strAttorneyOsmep} สำนักงานตั้งอยู่เลขที่ 120 หมู่ 3 ศูนย์ราชการเฉลิมพระเกียรติ 80 พรรษา 5 ธันวาคม 2550. (อาคารซี) ชั้น 2, 10, 11 ถนนแจ้งวัฒนะ แขวงทุ่งสองห้อง เขตหลักสี่ กรุงเทพ 10210 ซึ่งต่อไปในสัญญานี้จะเรียกว่า “ผู้เปิดเผยข้อมูล”
        กับ {result.Contract_Party_Name} โดย {result.CP_S_NAME} ตำแหน่ง {result.CP_S_POSITION} {strAttorney} 
        สำนักงานตั้งอยู่เลขที่ {result.OfficeLoc} ซึ่งต่อไปในสัญญานี้จะเรียกว่า “ผู้รับข้อมูล”
    </p>
   <p class='tab3 t-16'>คู่สัญญาได้ตกลงทำสัญญากันมีข้อความดังต่อไปนี้</p>

    <!-- NDA Clauses -->
    <!-- Purpose -->
   {purposeHtml}

   <p class='tab3 t-16'>โดยผู้รับข้อมูลประสงค์ให้ผู้ให้ข้อมูลเปิดเผยข้อมูลแก่ผู้รับข้อมูลอย่างเป็นความลับทั้งก่อน
หรือหลังจากวันที่สัญญาฉบับนี้มีผลใช้บังคับดังที่ระบุไว้ข้างต้น โดยผู้ให้ข้อมูลมีความจำเป็นต้องเปิดเผยข้อมูล
ที่เป็นความลับของผู้ให้ข้อมูล เพื่อผู้รับข้อมูลจะได้นำข้อมูลดังกล่าวไปประกอบการจัดทำ {result.Contract_Party_Name} ร่วมกัน โดยผู้ให้ข้อมูลประสงค์
ให้ผู้รับข้อมูลเก็บรักษาความลับไว้ภายใต้สัญญานี้
</p>

    <!-- Confidential Types -->
   <p class='tab3 t-16'><b>ข้อ ๒ ข้อมูลที่เป็นความลับ</b></p>
   <p class='tab3 t-16'>“ข้อมูลที่เป็นความลับ” หมายความว่า บรรดาข้อความเอกสารข้อมูลตลอดจนรายละเอียดทั้ง
ปวงที่เป็นของผู้ให้ข้อมูล รวมถึงที่อยู่ในความครอบครองหรือควบคุมดูแลของผู้ให้ข้อมูล และไม่เป็นที่รับรู้ของ
สาธารณชนโดยทั่วไปไม่ว่าจะในรูปแบบที่จับต้องได้หรือไม่ก็ตาม หรือสื่อแบบใดไม่ว่าจะถูกดัดแปลงแก้ไขโดย
ผู้รับข้อมูลหรือไม่ และไม่ว่าจะเปิดเผยเมื่อใดและอย่างไร ให้ถือว่าเป็นความลับโดยข้อมูลที่เป็นความลับอาจ
หมายความรวมถึง ข้อมูลเชิงกลยุทธ์ของผู้ให้ข้อมูล แผนธุรกิจ ข้อมูลทางการเงิน ข้อมูลลูกจ้าง ข้อมูลผู้
ประกอบการ และข้อมูลส่วนบุคคลที่ผู้ให้ข้อมูลได้เก็บ รวบรวม ใช้ ข้อมูลที่เป็นความลับที่ผู้ให้ข้อมูล หรือในนามของผู้ให้ข้อมูลที่เปิดเผยแก่ผู้รับข้อมูล ซึ่งหมายความรวมถึงข้อมูลที่ผู้ให้ข้อมูลให้แก่ผู้รับข้อมูล ดังนี้
    </p>
   <p class='tab3 t-16'>(ระบุประเภทของข้อมูลที่เป็นความลับที่นำส่งให้แก่กัน)</p>
    {conConfidentialHtml}

    <!-- Clause 3 -->
   <p class='tab3 t-16'><b>ข้อ ๓ การรักษาข้อมูลที่เป็นความลับ</b></p>
   <p class='tab3 t-16'>๓.๑ ผู้รับข้อมูลต้องรับผิดชอบรักษาข้อมูลที่เป็นความลับและเก็บข้อมูลความลับไว้โดยครบ
ถ้วนและอย่างเคร่งครัด ผู้รับข้อมูลจะต้องไม่เปิดเผยทำสำเนาหรือทำการอื่นใดทำนองเดียวกันแก่บุคคลอื่น
ไม่ว่าทั้งหมดหรือบางส่วน เว้นแต่ได้รับอนุญาตเป็นหนังสือจากผู้ให้ข้อมูล</p>

   <p class='tab3 t-16'>๓.๒ ผู้รับข้อมูลต้องใช้ข้อมูลที่เป็นความลับเพื่อการอันเกี่ยวกับหรือสัมพันธ์ 
กับการดำเนินงานที่มีอยู่ระหว่างผู้ให้ข้อมูลกับผู้รับข้อมูล โดยผู้รับข้อมูลต้องแจ้งให้ผู้ให้ข้อมูลทราบโดยทันที
ที่พบการใช้หรือการเปิดเผยข้อมูลที่เป็นความลับโดยไม่ได้รับอนุญาตหรือการละเมิดหรือฝ่าฝืนข้อกำหนด
ตามสัญญานี้ อีกทั้ง ผู้รับข้อมูลจะต้องให้ความร่วมมือกับผู้ให้ข้อมูลอย่างเต็มที่ในการเรียกคืนซึ่งการ 
ครอบครองข้อมูลที่เป็นความลับการป้องกันการใช้ข้อมูลที่เป็นความลับโดยไม่ได้รับอนุญาตและการระงับ 
ยับยั้งการเผยแพร่ข้อมูลที่เป็นความลับออกสู่สาธารณะ</p>
   <p class='tab3 t-16'>๓.๓ ผู้รับข้อมูลต้องจัดให้มีและคงไว้ซึ่งมาตรการรักษาความปลอดภัยสำหรับ 
การจัดเก็บและประมวลผลข้อมูลที่มีความเหมาะสมในมาตรการเชิงองค์กร มาตรการเชิงเทคนิค และ 
มาตรการเชิงกายภาพโดยคำนึงถึงลักษณะ ขอบเขต และวัตถุประสงค์ของการดำเนินการตามวัตถุประสงค์ 
ที่ของสัญญาฉบับนี้เป็นสำคัญ เพื่อป้องกันมิให้ข้อมูลที่เป็นความลับถูกนำไปใช้โดยมิได้รับอนุญาตหรือ 
ถูกเปิดเผยแก่บุคคลอื่น โดยผู้รับข้อมูลต้องใช้มาตรการการเก็บรักษาข้อมูลที่เป็นความลับในระดับเดียวกัน 
กับที่ผู้รับข้อมูลใช้กับข้อมูลที่เป็นความลับของตนเองซึ่งต้องไม่น้อยกว่าการดูแลที่สมควร</p>
   <p class='tab3 t-16'>๓.๔ ผู้รับข้อมูลต้องแจ้งให้บุคลากร พนักงาน ลูกจ้าง ที่ปรึกษาของผู้รับข้อมูล 
และ/หรือบุคคลภายนอกที่ต้องเกี่ยวข้องกับข้อมูลที่เป็นความลับนั้น ทราบถึงความเป็นความลับและ 
ข้อจำกัดสิทธิในการใช้และการเปิดเผยข้อมูลที่เป็นความลับ และผู้รับข้อมูลต้องดำเนินการให้บุคคลดังกล่าว 
ต้องผูกพันด้วยสัญญาหรือข้อตกลงเป็นหนังสือในการรักษาข้อมูลที่เป็นความลับ โดยมีข้อกำหนดเช่น 
เดียวกับหรือไม่น้อยกว่าข้อกำหนดและเงื่อนไขในสัญญาฉบับนี้ด้วย</p>
   <p class='tab3 t-16'>๓.๕ ข้อมูลที่เป็นความลับตามสัญญาฉบับนี้ไม่รวมไปถึงข้อมูลดังต่อไปนี้</p>
   <p class='tab4 t-16'>(๑) ข้อมูลที่ผู้ให้ข้อมูลเปิดเผยแก่สาธารณะ</p>
   <p class='tab4 t-16'>(๒) ข้อมูลที่ผู้รับข้อมูลทราบอยู่ก่อนที่ผู้ให้ข้อมูลจะเปิดเผยข้อมูลนั้น</p>
   <p class='tab4 t-16'>(๓) ข้อมูลที่มาจากการพัฒนาโดยอิสระของผู้รับข้อมูลเอง</p>
   <p class='tab4 t-16'>(๔) ข้อมูลที่ต้องเปิดเผยโดยกฎหมายหรือตามคำสั่งศาล ทั้งนี้ ผู้รับข้อมูลต้องมีหนัง
สือแจ้งผู้ให้ข้อมูลได้รับทราบถึงข้อกำหนดหรือคำสั่งดังกล่าว โดยแสดงเอกสารข้อกำหนด หมายศาลและ/
หรือหมายค้นอย่างเป็นทางการต่อผู้ให้ข้อมูลก่อนที่จะดำเนินการเปิดเผยข้อมูลดังกล่าว และในการเปิดเผย
ข้อมูลดังกล่าวผู้รับข้อมูลจะต้องดำเนินการตามขั้นตอนทางกฎหมายเพื่อขอให้คุ้มครองข้อมูลดังกล่าวไม่ให้
ถูกเปิดเผยต่อสาธารณะด้วย</p>
   <p class='tab4 t-16'>(๕) ผู้รับข้อมูลได้รับความยินยอมเป็นลายลักษณ์อักษรให้เปิดเผยข้อมูลจากผู้ให้ข้อมูล
ก่อนที่ผู้รับข้อมูลจะเปิด เผยข้อมูลนั้น</p>
   <p class='tab4 t-16'>(๖) ผู้รับข้อมูลได้รับข้อมูลที่เป็นความลับจากบุคคลที่สามที่ไม่อยู่ภายใต้ข้อกำหนดใน
เรื่องการรักษาความลับ หรือข้อจำกัดในเรื่องสิทธิ</p>
   <p class='tab3 t-16'>๓.๖ ผู้รับข้อมูลต้องไม่ทำซ้ำข้อมูลที่เป็นความลับแม้เพียงส่วนหนึ่งส่วนใดหรือทั้งหมด
เว้นแต่การทำซ้ำเพื่อการใช้ข้อมูลที่เป็นความลับให้บรรลุผลตามวัตถุประสงค์ที่กำหนดไว้ในสัญญานี้ และ
ไม่ทำวิศวกรรมย้อนกลับ หรือถอดรหัสข้อมูลที่เป็นความลับ ต้นแบบ หรือสิ่งอื่นใดที่บรรจุข้อมูลที่เป็น
ความลับ รวมทั้งไม่เคลื่อนย้าย พิมพ์ทับ หรือทำให้เสียรูปซึ่งสัญลักษณ์ที่แสดงเครื่องหมายสิทธิบัตร
อนุสิทธิบัตร ลิขสิทธิ์ เครื่องหมายการค้า ตราสัญลักษณ์ และเครื่องหมายอื่นใดที่แสดงกรรมสิทธิ์ของ
ต้นแบบหรือสำเนาของข้อมูลที่เป็นความลับที่ได้รับมาจากผู้ให้ข้อมูล</p>

    <!-- Clause 4 -->
   <p class='tab3 t-16'><b>ข้อ ๔ ทรัพย์สินทางปัญญา</b></p>
   <p class='tab3 t-16'>สัญญาฉบับนี้ไม่มีผลบังคับใช้เป็นการโอนสิทธิหรือการอนุญาตให้ใช้สิทธิ (ไม่ว่าโดยตรง
 หรือโดยอ้อม) ให้แก่ผู้รับข้อมูลที่ได้รับความลับซึ่งสิทธิบัตร ลิขสิทธิ์ การออกแบบ เครื่องหมายการค้าตรา
สัญลักษณ์ รูปประดิษฐ์อื่นใดชื่อทางการค้า ความลับทางการค้า ไม่ว่าจดทะเบียนไว้ตามกฎหมายหรือไม่
ก็ตามหรือสิทธิอื่น ๆ ของผู้ให้ข้อมูล ซึ่งอาจปรากฏอยู่หรือนำมาทำซ้ำไว้ในเอกสารข้อมูลที่เป็นความลับทั้งนี้
ผู้รับข้อมูลหรือบุคคลอื่นใดที่เกี่ยวข้องกับผู้รับข้อมูลและเกี่ยวข้องกับข้อมูลที่เป็นความลับดังกล่าวจะไม่ยื่น
ขอรับสิทธิและ/หรือขอจดทะเบียนเกี่ยวกับทรัพย์สินทางปัญญาใด ๆ ตลอดจนไม่นำไปใช้โดยไม่ได้รับ
การอนุญาตเป็นลายลักษณ์อักษรจากผู้ให้ข้อมูลเกี่ยวกับรายละเอียดข้อมูลที่เป็นความลับหรือส่วนหนึ่งส่วนใด
ของรายละเอียดดังกล่าว</p>

    <!-- Clause 5 -->
   <p class='tab3 t-16'><b>ข้อ ๕ การส่งคืน ลบ หรือการทำลายข้อมูลที่เป็นความลับ</b></p>
   <p class='tab3 t-16'>เมื่อการดำเนินงานที่มีอยู่ระหว่างผู้ให้ข้อมูลกับผู้รับข้อมูลเสร็จสิ้นลงตามวัตถุประสงค์ผู้รับ
ข้อมูลจะต้องส่งมอบข้อมูลที่เป็นความลับและสำเนาของข้อมูลที่เป็นความลับที่ผู้รับข้อมูลได้รับไว้คืนให้แก่ผู้ให้
ข้อมูล เว้นแต่ผู้ให้ข้อมูลเห็นว่าไม่ต้องนำส่งคืนแต่ต้องเลิกใช้ข้อมูลที่เป็นความลับ และทำการลบหรือทำลาย
ข้อมูลที่เป็นความลับทั้งถูกจัดเก็บไว้ในคอมพิวเตอร์หรืออุปกรณ์อื่นใดที่ใช้จัดเก็บข้อมูล (ถ้ามี) หรือดำเนิน
การอื่นตามที่ได้รับการแจ้งเป็นลายลักษณ์อักษรจากผู้ให้ข้อมูล</p>

    <!-- Clause 6 -->
   <p class='tab4 t-16'><b>ข้อ ๖ การชดใช้ค่าเสียหาย</b></p>
   <p class='tab4 t-16'>๖.๑ กรณีที่ผู้รับข้อมูล พนักงาน ลูกจ้าง ที่ปรึกษาของผู้รับข้อมูล และ/หรือบุคคลภาย
นอกที่ได้รับข้อมูลที่เป็นความลับจากผู้รับข้อมูลฝ่าฝืนข้อกำหนดตามสัญญานี้และก่อให้เกิดความเสียหายแก่ผู้
ให้ข้อมูล และ/หรือบุคคลอื่นผู้รับข้อมูลจะต้องชดใช้ค่าเสียหายให้แก่ผู้ให้ข้อมูล และ/หรือบุคคลที่ได้รับความ
เสียหายสำหรับความเสียหายเช่นว่านั้น ทั้งนี้ ผู้รับข้อมูลจะต้องแจ้งแก่ผู้ให้ข้อมูลทราบเป็นลายลักษณ์อักษร
ภายใน ๗ วันนับตั้งแต่มีการละเมิดข้อมูลที่เป็นความลับเกิดขึ้น</p>

   <p class='tab4 t-16'>๖.๒ ผู้รับข้อมูลรับทราบว่าการเปิดเผยหรือการใช้ข้อมูลที่เป็นความลับโดยฝ่าฝืนข้อ
กำหนดตามสัญญานี้จะก่อให้เกิดความเสียหายแก่ผู้ให้ข้อมูลในจำนวนที่ไม่สามารถประเมินได้ดังนั้นผู้รับ
ข้อมูลยินยอมให้ผู้ให้ข้อมูลใช้สิทธิที่จะร้องขอต่อศาลเพื่อให้มีคำสั่งให้ผู้รับข้อมูลหยุดการกระทำใดๆที่เป็น
การฝ่าฝืนข้อกำหนดตามสัญญานี้ และ/หรือใช้วิธีคุ้มครองชั่วคราวใดๆ ตามที่ผู้ให้ข้อมูลเห็นว่าเหมาะสม
ได้โดยผู้รับข้อมูลจะเป็นผู้รับผิดชอบค่าใช้จ่ายต่าง ๆ ที่เกิดขึ้นทั้งหมดจากการดำเนินการดังกล่าว</p>
   <p class='tab3 t-16'>๖.๓ กรณีที่ผู้ให้ข้อมูลสงสัยว่าผู้รับข้อมูลฝ่าฝืนข้อกำหนดตามสัญญานี้ ผู้รับข้อมูล 
จะต้องเป็นฝ่ายพิสูจน์ว่าผู้รับข้อมูลไม่ได้ฝ่าฝืนข้อกำหนดตามสัญญานี้</p>

    <!-- Clause 7 -->
   <p class='tab3 t-16'><b>ข้อ ๗ ระยะเวลาตามสัญญา</b></p>
   <p class='tab3 t-16'>สัญญานี้มีผลบังคับใช้นับตั้งแต่วันที่ทำสัญญานี้ โดยมีกำหนดระยะเวลาทั้งสิ้น {result.EnforcePeriods??"-"} ปี นับตั้งแต่
วันที่ทำสัญญาฉบับนี้</p>
   <p class='tab3 t-16'>เมื่อครบกำหนดระยะเวลาตามวรรคหนึ่ง หรือเมื่อมีการบอกเลิกสัญญา หรือผู้ให้ข้อมูล
ได้แจ้งให้ผู้รับข้อมูลดำเนินการทำลายข้อมูลดังกล่าว ผู้รับข้อมูลจะต้องดำเนินการทำลายข้อมูล ภายใน ๗
วันนับแต่ได้รับหนังสือร้องขอจากผู้ให้ข้อมูล ทั้งนี้ ผู้รับข้อมูลจะต้องไม่มีการสงวนไว้ซึ่งสำเนาใด ๆ</p>

    <!-- Clause 8 -->
   <p class='tab3 t-16'>ข้อ ๘ ข้อตกลงอื่น ๆ</p>
   <p class='tab3 t-16'>๘.๑ ในกรณีที่มีเหตุจำเป็นต้องมีการเปลี่ยนแปลงแก้ไขสัญญานี้ ให้ทำเป็นลายลักษณ์อักษร
และลงนามโดยคู่สัญญาหรือผู้มีอำนาจลงนามผูกพันนิติบุคคลและประทับตราสำคัญของนิติบุคคล (ถ้ามี) ของคู่สัญญา แล้วแต่กรณี</p>
   <p class='tab3 t-16'>๘.๒ กรณีที่ผู้รับข้อมูลได้โอนกิจการ รวมกิจการ หรือควบกิจการ หรือดำเนินการอื่น ๆ ใน
ลักษณะที่มีการเปลี่ยนแปลงของวัตถุประสงค์ในการดำเนินกิจการของผู้รับข้อมูลผู้รับข้อมูลจะต้องแจ้งให้ผู้ให้
ข้อมูลทราบภายใน ๕ วันทำการ นับแต่ได้เกิดเหตุดังกล่าวขึ้น</p>

    <!-- Clause 9 -->
   <p class='tab3 t-16'>ข้อ ๙ การบังคับใช้</p>
   <p class='tab3 t-16'>๙.๑ ในกรณีที่ปรากฏภายหลังว่าส่วนใดส่วนหนึ่งในสัญญาฉบับนี้เป็นโมฆะให้ถือว่า 
ข้อกำหนดส่วนที่เป็นโมฆะไม่มีผลบังคับในสัญญานี้ และข้อกำหนดที่เหลืออยู่ในสัญญาฉบับนี้ยังคงใช้บังคับ 
และมีผลอยู่อย่างสมบูรณ์</p>
   <p class='tab3 t-16'>๙.๒ สัญญาฉบับนี้อยู่ภายใต้การบังคับและตีความตามกฎหมายของประเทศไทย ให้ศาลของ
ประเทศไทยมีอำนาจในกรณีที่มีข้อพิพาทใด ๆ อันเกิดขึ้นจากสัญญาฉบับนี้</p>
   <p class='tab3 t-16'>สัญญานี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน คู่สัญญาได้อ่าน และเข้าใจข้อความใน
สัญญาโดยละเอียดตลอดแล้ว เห็นว่าตรงตามเจตนารมณ์ทุกประการ จึงได้ลงลายมือชื่อพร้อมทั้งประทับตรา
สำคัญผูกพันนิติบุคคล (ถ้ามี) ไว้เป็นสำคัญ ณ วัน เดือน ปี ที่ระบุข้างต้น และคู่สัญญาต่างฝ่ายต่าง
ยึดถือไว้ฝ่ายละหนึ่งฉบับ</p>

    <!-- Signature Table -->
   </br>
</br>
{signatoryTableHtml}
</body>
</html>
";


        return html;
    }
    #endregion  4.1.1.2.7.สัญญาการรักษาข้อมูลที่เป็นความลับ
}
