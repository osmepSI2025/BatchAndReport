using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class WordEContract_PersernalProcessService
{
    private readonly WordServiceSetting _w;

    public WordEContract_PersernalProcessService(WordServiceSetting ws)
    {
        _w = ws;
    }
    #region 4.1.1.2.4.บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล
    public byte[] OnGetWordContact_PersernalProcessService()
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
            body.AppendChild(WordServiceSetting.CenteredBoldParagraph("ข้ออตกลงการประมวลผลข้อมูลส่วนบุคคล", "44"));
            body.AppendChild(WordServiceSetting.CenteredBoldParagraph("(Data Processing Agreement)", "44"));
            body.AppendChild(WordServiceSetting.CenteredBoldParagraph("โครงการ.....(ระบุชื่อบันทึกข้อตกลงความร่วมมือหรือสัญญาฉบับหลัก)....", "44"));
        
            body.AppendChild(WordServiceSetting.CenteredBoldParagraph("ระหว่าง", "32"));
  
            body.AppendChild(WordServiceSetting.CenteredBoldParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม กับ ……..(ชื่อคู่สัญญา)…..….", "32"));
  body.AppendChild(WordServiceSetting.CenteredBoldParagraph("---------------------------------", "36"));



  // --- 3. Main contract body ---
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อตกลงการประมวลผลข้อมูลส่วนบุคคล (“ข้อตกลง”) ฉบับนี้ทำขึ้น เมื่อวันที่.... (ระบุวันที่ลงนามในข้อตกลง)....... ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม", null, "32"));
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("โดยที่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สสว.” ฝ่ายหนึ่ง ได้ตกลงใน...........(ระบุชื่อบันทึกข้อตกลงความร่วมมือ/สัญญาหลัก)................ สัญญาเลขที่ .......... (ระบุเลขที่บันทึกข้อตกลงความร่วมมือ/สัญญาหลัก)................. ฉบับลงวันที่ ..... (ระบุวันที่ลงนามข้อตกลงความร่วมมือหรือวันทำสัญญาหลัก).......... ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “(บันทึกความร่วมมือ/สัญญา)” กับ ........(ระบุชื่อคู่สัญญา)........ ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “.....(ระบุชื่อเรียกคู่สัญญา......” อีกฝ่ายหนึ่ง", null, "32"));
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ตามที่ (ระบุชื่อบันทึกความร่วมมือ/สัญญาหลัก) ดังกล่าวกำหนดให้ สสว. มีหน้าที่และความรับผิดชอบในส่วนของการ.......(ระบุขอบเขต สิทธิ หน้าที่ของ สสว. ตามบันทึกความร่วมมือ/สัญญาหลัก)...... ซึ่งในการดำเนินการดังกล่าวประกอบด้วยการมอบหมายหรือแต่งตั้งให้...... (ระบุชื่อคู่สัญญา)......เป็นผู้ดำเนินการกระบวนการเก็บรวบรวม ใช้ หรือเปิดเผย (“ประมวลผล”) ข้อมูลส่วนบุคคลแทนหรือในนามของ สสว. ", null, "32"));
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สสว. ในฐานะผู้ควบคุมข้อมูลส่วนบุคคลเป็นผู้มีอำนาจตัดสินใจ กำหนดรูปแบบและกำหนดวัตถุประสงค์ในการประมวลผลข้อมูลส่วนบุคคล ได้.....(มอบหมาย/แต่งตั้ง/จ้าง/อื่น ๆ).....ให้.....(ระบุชื่อคู่สัญญา).......ในฐานะผู้ประมวลผลข้อมูลส่วนบุคคล ดำเนินการเพื่อวัตถุประสงค์ดังต่อไปนี้", null, "32"));
   
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("1. ..... (ระบุวัตถุประสงค์ที่ สสว. มอบหมายให้คู่สัญญาดำเนินการเกี่ยวกับข้อมูลส่วนบุคคล เช่น เพื่อการรับจ้างทำระบบยืนยันตัวตน เพื่อการรับทำ Survey เพื่อการลงทะเบียนผู้เข้าร่วมงานสัมมนา เพื่อการรับจ้างพิมพ์บัตรพนักงาน เพื่อการรับส่งเอกสาร เป็นต้น).........", null, "32"));
  
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("2. ............................................................โดยข้อมูลส่วนบุคคลที่ สสว. มอบหมาย.....(มอบหมาย/แต่งตั้ง/จ้าง/อื่น ๆ).....ให้.... (ระบุชื่อคู่สัญญา).....ประมวลผล ประกอบด้วย", null, "32"));
  
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("1. ..... (ระบุรายการข้อมูลส่วนบุคคลที่ สสว. มอบหมาย/เปิดเผยให้คู่สัญญาประมวลผล เช่น ชื่อ นามสกุลของเจ้าหน้าที่ เบอร์โทรศัพท์ ข้อมูลผู้ใช้งานแอปพลิเคชั่นทางรัฐ รายชื่อผู้เข้าร่วมงานสัมมนา เป็นต้น).........", null, "32"));
   
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("2. ............................................................", null, "32"));
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ด้วยเหตุนี้ ทั้งสองฝ่ายจึงตกลงจัดทำข้อตกลงฉบับนี้ และให้ถือข้อตกลงฉบับนี้เป็นส่วนหนึ่งของ....(ระบุชื่อบันทึกข้อตกลงความร่วมมือ/สัญญาหลัก)....เพื่อเป็นหลักฐานการควบคุมดูแลการประมวลผลข้อมูลส่วนบุคคลที่ สสว. มอบหมายหรือแต่งตั้งให้......... (ระบุชื่อคู่สัญญา)........... ดำเนินการ อันเนื่องมาจากการดำเนินการตามหน้าที่และความรับผิดชอบตาม....(ระบุชื่อบันทึกข้อตกลงความร่วมมือ/สัญญาหลัก)....ฉบับลงวันที่ ..... (ระบุวันที่ลงนามข้อตกลงความร่วมมือหรือวันทำสัญญาหลัก)......... และเพื่อดำเนินการให้เป็นไปตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒ และกฎหมายอื่น ๆ ที่ออกตามความในพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒ ซึ่งต่อไปในข้อตกลงฉบับนี้ รวมเรียกว่า “กฎหมายคุ้มครองข้อมูลส่วนบุคคล” ทั้งที่มีผลใช้บังคับอยู่ ณ วันทำข้อตกลงฉบับนี้และที่จะมีการเพิ่มเติมหรือแก้ไขเปลี่ยนแปลงในภายหลัง โดยมีรายละเอียดดังนี้ ", null, "32"));
   
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("1. ........ (ระบุชื่อคู่สัญญา)........ รับทราบว่า ข้อมูลส่วนบุคคล หมายถึง ข้อมูลเกี่ยวกับบุคคลธรรมดาซึ่งทำให้สามารถระบุตัวบุคคลนั้นได้ไม่ว่าทางตรงหรือทางอ้อม โดย........ (ระบุชื่อคู่สัญญา)........ จะดำเนินการ ตามที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลกำหนด เพื่อคุ้มครองให้การประมวลผลข้อมูลส่วนบุคคลเป็นไปอย่างเหมาะสมและถูกต้องตามกฎหมาย", null, "32"));
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("โดยในการดำเนินการตามข้อตกลงนี้ ......... (ระบุชื่อคู่สัญญา)............จะประมวลผลข้อมูลส่วนบุคคลเมื่อได้รับคำสั่งที่เป็นลายลักษณ์อักษรจาก สสว. แล้วเท่านั้น ทั้งนี้ เพื่อให้ปราศจากข้อสงสัย การดำเนินการประมวลผลข้อมูลส่วนบุคคลโดย....... (ระบุชื่อคู่สัญญา)............ตามหน้าที่และความรับผิดชอบตาม....(ระบุชื่อบันทึกข้อตกลงความร่วมมือ/สัญญาหลัก)....ถือเป็นการได้รับคำสั่งที่เป็นลายลักษณ์อักษรจาก สสว. แล้ว", null, "32"));
   
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("2. ......... (ระบุชื่อคู่สัญญา)............จะกำหนดให้การเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้ถูกจำกัดเฉพาะเจ้าหน้าที่ และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ได้รับมอบหมาย มีหน้าที่เกี่ยวข้องหรือมีความจำเป็นในการเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้เท่านั้น และจะดำเนินการเพื่อให้พนักงาน และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ได้รับมอบหมายจาก......... (ระบุชื่อคู่สัญญา)............ทำการประมวลผลและรักษาความลับของข้อมูลส่วนบุคคลด้วยมาตรฐานเดียวกัน", null, "32"));
  
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3. ......... (ระบุชื่อคู่สัญญา)............จะควบคุมดูแลให้เจ้าหน้าที่ และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ปฏิบัติหน้าที่ในการประมวลผลข้อมูลส่วนบุคคล ปฏิบัติตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลอย่างเคร่งครัด และดำเนินการประมวลผลข้อมูลส่วนบุคคลตามวัตถุประสงค์ของการดำเนินการตามข้อตกลงฉบับนี้เท่านั้น โดยจะไม่ทำซ้ำ คัดลอก ทำสำเนา บันทึกภาพข้อมูลส่วนบุคคลไม่ว่าทั้งหมดหรือแต่บางส่วนเป็นอันขาด เว้นแต่เป็นไปตามเงื่อนไขของบันทึกความร่วมมือหรือสัญญา หรือกฎหมายที่เกี่ยวข้องจะระบุหรือบัญญัติไว้เป็นประการอื่น", null, "32"));
   
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("4. ......... (ระบุชื่อคู่สัญญา).............จะดำเนินการเพื่อช่วยเหลือหรือสนับสนุน สสว. ในการตอบสนองต่อคำร้องที่เจ้าของข้อมูลส่วนบุคคลแจ้งต่อ สสว. อันเป็นการใช้สิทธิของเจ้าของข้อมูลส่วนบุคคลตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลในส่วนที่เกี่ยวข้องกับการประมวลผลข้อมูลส่วนบุคคลในขอบเขตของข้อตกลงฉบับนี้ ", null, "32"));
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("อย่างไรก็ดี ในกรณีที่เจ้าของข้อมูลส่วนบุคคลยื่นคำร้องขอใช้สิทธิดังกล่าวต่อ......... (ระบุชื่อคู่สัญญา).............โดยตรง ......... ", null, "32"));
  body.AppendChild(WordServiceSetting.JustifiedParagraph("(ระบุชื่อคู่สัญญา)............จะดำเนินการแจ้งและส่งคำร้องดังกล่าวให้แก่ สสว. ทันที โดย......... ", "32"));
  body.AppendChild(WordServiceSetting.JustifiedParagraph("(ระบุชื่อคู่สัญญา)........จะไม่เป็นผู้ตอบสนองต่อคำร้องดังกล่าว เว้นแต่ สสว. จะได้มอบหมายให้......... ", "32"));
  body.AppendChild(WordServiceSetting.JustifiedParagraph("(ระบุชื่อคู่สัญญา).......ดำเนินการเฉพาะเรื่องที่เกี่ยวข้องกับคำร้องดังกล่าว", "32"));

  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("5. ......... (ระบุชื่อคู่สัญญา)............จะจัดทำและเก็บรักษาบันทึกรายการของกิจกรรมการประมวลผลข้อมูลส่วนบุคคล (Record of Processing Activities) ทั้งหมดที่......... (ระบุชื่อคู่สัญญา)............ประมวลผลในขอบเขตของข้อตกลงฉบับนี้ และจะดำเนินการส่งมอบบันทึกรายการดังกล่าวให้แก่ สสว. ทุก.....(ระบุความถี่ของการส่งมอบบันทึกรายการ เช่น ทุกสัปดาห์หรือทุกเดือน).... และ/หรือทันทีที่ สสว. ร้องขอ", null, "32"));
   
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("6. ......... (ระบุชื่อคู่สัญญา)............จะจัดให้มีและคงไว้ซึ่งมาตรการรักษาความปลอดภัยสำหรับการประมวลผลข้อมูลที่มีความเหมาะสมทั้งในเชิงองค์กรและเชิงเทคนิคตามที่คณะกรรมการคุ้มครองข้อมูลส่วนบุคคลได้ประกาศกำหนดและ/หรือตามมาตรฐานสากล โดยคำนึงถึงลักษณะ ขอบเขต และวัตถุประสงค์ของการประมวลผลข้อมูลตามที่กำหนดในข้อตกลงฉบับนี้เป็นสำคัญ เพื่อคุ้มครองข้อมูลส่วนบุคคลจากความเสี่ยงอันเกี่ยวเนื่องกับการประมวลผลข้อมูลส่วนบุคคล เช่น ความเสียหายอันเกิดจากการละเมิด อุบัติเหตุ การลบ ทำลาย สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้ เปิดเผยหรือโอนข้อมูลส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย เป็นต้น", null, "32"));
   
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("7. เว้นแต่กฎหมายที่เกี่ยวข้องจะบัญญัติไว้เป็นประการอื่น ......... (ระบุชื่อคู่สัญญา)............จะทำการลบหรือทำลายข้อมูลส่วนบุคคลที่ทำการประมวลผลภายใต้ข้อตกลงฉบับนี้ภายใน....(ระบุจำนวนวันที่จะทำการลบทำลายข้อมูล).....วัน นับแต่วันที่ดำเนินการประมวลผลเสร็จสิ้น หรือวันที่ สสว. และ ......... (ระบุชื่อคู่สัญญา)............ได้ตกลงเป็นลายลักษณ์อักษรให้ยกเลิก....(ระบุชื่อบันทึกข้อตกลงความร่วมมือ/สัญญาหลัก)....แล้วแต่กรณีใดจะเกิดขึ้นก่อน ", null, "32"));
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("นอกจากนี้ ในกรณีปรากฏว่า......... (ระบุชื่อคู่สัญญา)........หมดความจำเป็นจะต้องเก็บรักษาข้อมูลส่วนบุคคลตามข้อตกลงฉบับนี้ก่อนสิ้นระยะเวลาตามวรรคหนึ่ง .......... (ระบุชื่อคู่สัญญา)............จะทำการลบหรือทำลายข้อมูลส่วนบุคคลตามข้อตกลงฉบับนี้ทันที", null, "32"));
   
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("8. กรณีที่......... (ระบุชื่อคู่สัญญา)............พบพฤติการณ์ใด ๆ ที่มีลักษณะที่กระทบต่อการรักษาความปลอดภัยของข้อมูลส่วนบุคคลที่......... (ระบุชื่อคู่สัญญา).............ประมวลผลภายใต้ข้อตกลงฉบับนี้ ซึ่งอาจก่อให้เกิดความเสียหายจากการละเมิด อุบัติเหตุ การลบ ทำลาย สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้ เปิดเผยหรือโอนข้อมูลส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย แล้ว......... (ระบุชื่อคู่สัญญา)............จะดำเนินการแจ้งให้ สสว. ทราบโดยทันทีภายในเวลาไม่เกิน....(ระบุเวลาเป็นหน่วยชั่วโมงที่คู่สัญญาต้องแจ้งเหตุแก่ สสว. เช่น ภายใน 24 ชั่วโมงหรือ 48 ชั่วโมง ทั้งนี้ไม่ควรเกิน 48 ชั่วโมงเนื่องจาก สสว. ในฐานะผู้ควบคุมข้อมูลส่วนบุคคลมีหน้าที่ต้องแจ้งเหตุดังกล่าวแก่คณะกรรมการคุ้มครองข้อมูลส่วนบุคคลภายใน 72 ชั่วโมง).... ชั่วโมง", null, "32"));
  
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("9. การแจ้งถึงเหตุการละเมิดข้อมูลส่วนบุคคลที่เกิดขึ้นภายใต้ข้อตกลงนี้......... (ระบุชื่อคู่สัญญา)............จะใช้มาตรการตามที่เห็นสมควรในการระบุถึงสาเหตุของการละเมิด และป้องกันปัญหาดังกล่าวมิให้เกิดซ้ำ และจะให้ข้อมูลแก่ สสว. ภายใต้ขอบเขตที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลได้กำหนด ดังต่อไปนี้", null, "32"));
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("- รายละเอียดของลักษณะและผลกระทบที่อาจเกิดขึ้นของการละเมิด", null, "32"));
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("- มาตรการที่ถูกใช้เพื่อลดผลกระทบของการละเมิด", null, "32"));
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("- ข้อมูลอื่น ๆ เกี่ยวข้องกับการละเมิด", null, "32"));
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("10. หน้าที่และความรับผิดของ......... (ระบุชื่อคู่สัญญา)............ในการปฏิบัติตามข้อตกลงจะสิ้นสุดลงนับแต่วันที่ปฏิบัติงานที่ตกลงเสร็จสิ้น หรือ วันที่......... (ระบุชื่อคู่สัญญา)............และ สสว. ได้ตกลงเป็นลายลักษณ์อักษรให้ยกเลิก....(ระบุชื่อบันทึกข้อตกลงความร่วมมือ/สัญญาหลัก)....แล้วแต่กรณีใดจะเกิดขึ้นก่อน อย่างไรก็ดี การสิ้นผลลงของข้อตกลงนี้ ไม่กระทบต่อหน้าที่ของ......... (ระบุชื่อคู่สัญญา)............ในการลบหรือทำลายข้อมูลส่วนบุคคลตามที่ได้กำหนดในข้อ 7 ของข้อตกลงฉบับนี้", null, "32"));
  body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ทั้งสองฝ่ายได้อ่านและเข้าใจข้อความโดยละเอียดตลอดแล้ว เพื่อเป็นหลักฐานแห่งการนี้ ทั้งสองฝ่ายจึงได้ลงนามไว้เป็นหลักฐานต่อหน้าพยาน ณ วัน เดือน ปี ที่ระบุข้างต้น", null, "32"));


  
          body.AppendChild(WordServiceSetting.EmptyParagraph());
            body.AppendChild(WordServiceSetting.EmptyParagraph());
            body.AppendChild(WordServiceSetting.RightParagraph("ลงชื่อ.......................................ลงชื่อ......................................."));
            body.AppendChild(WordServiceSetting.RightParagraph("(................................................................................)"));
            body.AppendChild(WordServiceSetting.RightParagraph("(.............................ชื่อเต็มหน่วยงาน...................................)"));

            body.AppendChild(WordServiceSetting.RightParagraph("ลงชื่อ......................................................................พยาน"));
            body.AppendChild(WordServiceSetting.RightParagraph("(...............................................................................)"));
            body.AppendChild(WordServiceSetting.RightParagraph("ลงชื่อ......................................................................พยาน"));
            body.AppendChild(WordServiceSetting.RightParagraph("(...............................................................................)"));

            // --- 6. Signature lines ---
            body.AppendChild(WordServiceSetting.EmptyParagraph());


            // --- 7. Add header/footer if needed ---
            WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);
        }
        stream.Position = 0;
        return stream.ToArray();
    }
    #endregion 4.1.1.2.4.บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล
}
