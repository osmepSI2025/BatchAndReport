using BatchAndReport.DAO;
using BatchAndReport.Entities;
using BatchAndReport.Models;
using BatchAndReport.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;


public class WordEContract_BorrowMoneyService
{
    private readonly WordServiceSetting _w;

    public WordEContract_BorrowMoneyService(WordServiceSetting ws)
    {
        _w = ws;
    }
    #region  สสว. สัญญาเงินกู้ยืม โครงการพลิกฟื้นวิสาห
    public byte[] OnGetWordContact_orrowMoney()
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
            if (System.IO.File.Exists(imagePath))
            {
                var imagePart = mainPart.AddImagePart(ImagePartType.Png);
                using (var imgStream = new FileStream(imagePath, FileMode.Open))
                {
                    imagePart.FeedData(imgStream);
                }
                var element = WordServiceSetting.CreateImage(mainPart.GetIdOfPart(imagePart), 240, 80);
                var logoPara = new Paragraph(
                    new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                    element
                );
                body.AppendChild(logoPara);
            }

            // 2. Document title and subtitle
            body.AppendChild(WordServiceSetting.EmptyParagraph());
            body.AppendChild(WordServiceSetting.RightParagraph("ทะเบียนลูกค้า ............................"));
            body.AppendChild(WordServiceSetting.RightParagraph("เลขที่สัญญา ............................"));
            body.AppendChild(WordServiceSetting.EmptyParagraph());
            body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("สัญญากู้ยืมเงิน", "FF0000")); // Blue
            body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("โครงการเงินทุนพลิกฟื้นวิสาหกิจขนาดย่อม", "FF0000")); // Red
            body.AppendChild(WordServiceSetting.RightParagraph("ทำที่ ธนาคารพัฒนาวิสาหกิจขนาดกลางและขนาดย่อมแห่งประเทศไทย"));
            body.AppendChild(WordServiceSetting.RightParagraph("สำนักงานใหญ่/สาขา.........................................................."));

            // 3. Fillable lines (using underlines)
            // body.AppendChild(EmptyParagraph());
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้าพเจ้า ...................................................................................................................."));
            body.AppendChild(WordServiceSetting.JustifiedParagraph("อายุ ......... ปี สัญชาติ .................. สำนักงาน/บ้านตั้งอยู่เลขที่.................. อาคาร..........................................."));
            body.AppendChild(WordServiceSetting.JustifiedParagraph("หมู่ที่...........ตรอก/ซอย..........................ถนน...........................ตำบล/แขวง.................. ..................."));
            body.AppendChild(WordServiceSetting.JustifiedParagraph("เขต/อำเภอ................... จังหวัด.................ทะเบียนนิติบุคคลเลขที่/เลขประจำตัวประชาชนที่............................................"));
            body.AppendChild(WordServiceSetting.JustifiedParagraph("จดทะเบียนเป็นนิติบุคคลเมื่อวันที่..................................ซึ่งต่อไปนี้จะเรียกบุคคลผู้มีนามตามที่ระบุข้างต้นทั้งหมดว่า \"ผู้กู้\"ได้ทำสัญญาฉบับนี้ให้ไว้แก่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม" +
             "ซึ่งต่อไปนี้จะเรียกว่า \"ผู้ให้กู้\" โดยมีสาระสำคัญดังนี้"));

            // 4. Main body (sample)
            //  body.AppendChild(EmptyParagraph());
            body.AppendChild(WordServiceSetting.NormalParagraph("ซึ่งต่อไปนี้จะเรียกบุคคลผู้มีนามตามที่ระบุข้างต้นทั้งหมดว่า \"ผู้รับการอุดหนุน\" ได้ทำสัญญาฉบับนี้ให้ไว้แก่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม " +
                "ซึ่งต่อไปนี้จะเรียกว่า \"ผู้ให้การอุดหนุน\" โดยมีสาระสำคัญดังนี้", null, "28"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 1. วัตถุประสงค์และวงเงินกู้", null, "28", true));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("โดยผู้กู้ได้กู้เงินจากผู้ให้กู้เป็นจำนวนเงิน.....................................บาท(....................................................)เ" +
             "พื่อนำไปใช้จ่ายเป็นเงินทุนหมุนเวียน"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2TabsColor("โดยไม่นำเงินที่กู้ยืมไปชำระหนี้ที่มีอยู่ก่อนยื่นคำขอกู้ยืมเงิน", null, "FFF0000"));

            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("กำหนดชำระเงินกู้เสร็จสิ้นภายใน...........ปี...........เดือน  โดยมีระยะเวลาปลอดเงินต้น................เดือน"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 2. การเบิกจ่ายเงินกู้", null, "28", true));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ผู้ให้กู้จะจ่ายเงินกู้แก่ผู้กู้ตามเงื่อนไขการใช้เงินกู้ในข้อ 1.และตามรายละเอียดการใช้เงินกู้ ซึ่งผู้กู้ได้แจ้งไว้ในคำขอสินเชื่อและเอกสารแนบท้ายคำขอสินเชื่อโดยถือเป็นส่วนหนึ่งของสัญญากู้เงินฉบับนี้ด้วย" +
             "หากปรากฏว่ารายการขอเบิกเงินกู้งวดใดไม่เป็นไปตามเงื่อนไขและรายละเอียดดังกล่าว " +
             "เป็นสิทธิของผู้ให้กู้แต่ฝ่ายเดียวที่จะพิจารณาไม่ให้เบิกเงินกู้ก็ได้"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2TabsColor("โดยผู้ให้กู้จะจ่ายเงินกู้ให้ผู้กู้ด้วยการนำเงินหรือโอนเงินเข้าบัญชีที่ ธนาคารกรุงไทย จำกัด (มหาชน)" +
             "\r\nสาขา..............................................ชื่อบัญชี................................................................................ซึ่งเป็นบัญชีของผู้กู้" +
             "\r\nเลขที่บัญชี...................................จำนวนเงิน...........................บาท (.............................................)" +
             "และให้ถือว่าผู้กู้ได้รับเงินกู้ตามสัญญานี้ไปจากผู้ให้กู้แล้ว ในวันที่เงินเข้าบัญชีของผู้กู้ดังกล่าว"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2TabsColor("ทั้งนี้ ผู้กู้ยินยอมให้ผู้ให้กู้ หรือ ธนาคารพัฒนาวิสาหกิจขนาดกลางและขนาดย่อมแห่งประเทศไทย" +
             "ซึ่งกระทำการแทนผู้ให้กู้ หักเงินจากจำนวนเงินกู้ที่ผู้กู้ขอเบิกจากผู้ให้กู้เป็นค่าวิเคราะห์โครงการ ค่าอากรแสตมป์ ค่าใช้จ่ายหรือค่าธรรมเนียมในการโอนเงินเข้าบัญชีของผู้กู้ซึ่งธนาคารกรุงไทย จำกัด (มหาชน)" +
             " เรียกเก็บตามระเบียบของธนาคาร โดยไม่ต้องบอกกล่าวหรือแจ้งให้ผู้กู้ทราบ" +
             "โดยให้ถือว่าผู้กู้ได้รับเงินกู้ตามจำนวนที่เบิกไปครบถ้วนแล้วและสละสิทธิ์ที่จะเรียกร้องอย่างใด ๆ" +
             "ต่อผู้ให้กู้และหรือธนาคารพัฒนาวิสาหกิจขนาดกลางและขนาดย่อมแห่งประเทศไทย ที่ดำเนินการแทนตามที่ได้รับมอบหมายจากผู้ให้กู้", null, "FF0000"));

            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 3. ดอกเบี้ย", null, "28", true));

            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.1 การกู้ยืมเงินตามสัญญากู้เงินนี้ ไม่มีดอกเบี้ยเงินกู้"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.2 กรณีที่ผู้กู้ผิดเงื่อนไขการผ่อนชำระหนี้ และ/หรือไม่สามารถชำระหนี้เงินต้นคืนให้แก่ผู้ให้กู้ได้ครบถ้วนเมื่อครบกำหนดตามสัญญา" +
             "ผู้กู้และผู้ให้กู้ตกลงกันให้เป็นสิทธิของผู้ให้กู้ที่จะปรับอัตราดอกเบี้ยระหว่างผิดนัดการชำระหนี้ได้ในอัตราร้อยละ 15 ต่อปีโดยไม่ต้องบอกกล่าวผู้กู้" +
             "และ/หรือ ดำเนินการปรับโครงสร้างหนี้ให้แก่ผู้กู้ได้โดยผู้ให้กู้มีสิทธิที่จะคิดดอกเบี้ยจากผู้กู้ได้ในอัตราร้อยละ 15" +
             "ต่อปีจนกว่าจะชำระหนี้ให้แก่ผู้ให้กู้จนเสร็จสิ้น ตลอดจนดำเนินการใดๆ ได้ตามขอบเขตของประมวลกฎหมายแพ่งและพาณิชย์"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 4. การชำระคืนเงินต้นหรือชำระหนี้อื่นใด ให้แก่ผู้ให้กู้", null, "28", true));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("4.1 ผู้กู้ตกลงผ่อนชำระเงินต้นคืนให้แก่ผู้ให้กู้เป็นรายเดือนไม่น้อยกว่าเดือนละ..................... บาท" +
             " (..............................................................)" +
             " โดยชำระภายในวันที่.............ของทุกเดือน เริ่มตั้งแต่เดือน..........................พ.ศ. ........ เป็นต้นไป"));

            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("4.2 การชำระเงินตามข้อ 4.1  ผู้กู้ตกลงจะนำเงินเข้าบัญชีเงินฝากของผู้กู้ที่เปิดบัญชีไว้กับธนาคารกรุงไทย จำกัด (มหาชน) ตาม" +
             "ข้อ 2. โดยผู้กู้ยินยอมให้ ธนาคารพัฒนาวิสาหกิจขนาดกลางและขนาดย่อมแห่งประเทศไทยซึ่งดำเนินการแทนผู้ให้กู้ ในการแจ้งธนาคารเจ้าของบัญชีตาม" +
             "ข้อ 2. ให้หักเงินจากบัญชีของผู้กู้ดังกล่าวแล้วเพื่อชำระคืนเงินกู้แก่ผู้ให้กู้ในแต่ละงวดเดือน พร้อมทำการโอนเงินที่หักจากบัญชีของผู้กู้เพื่อนำเข้าบัญชีของผู้ให้กู้ที่เปิดบัญชี ไว้กับธนาคารกรุงไทย จำกัด (มหาชน)" +
             "สาขา............................................  บัญชีออมทรัพย์  ชื่อบัญชี" +
             "โครงการเงินทุนพลิกฟื้นวิสาหกิจขนาดย่อม เลขที่บัญชี............................................ เพื่อชำระหนี้คืนเงินกู้แก่ผู้ให้กู้ตามข้อตกลงในแต่ละงวดเดือน" +
             " เมื่อผู้ให้กู้ได้รับชำระเงินกู้คืนในแต่ละงวดแล้วจะออกใบเสร็จรับเงินให้แก่ผู้กู้ไว้เป็นหลักฐานต่อไป โดยผู้กู้ตกลงยินยอมให้หักเงินค่าธรรมเนียม\r\nในการโอนเงินชำระหนี้เงินกู้หรือค่าธรรมเนียมใด ๆ" +
             "ที่ธนาคารเจ้าของบัญชีเรียกเก็บในการโอนเงินจากบัญชีของผู้กู้ไปยังบัญชีเงินฝากของผู้ให้กู้ตามข้อ 4.2 ข้างต้นด้วย"));

            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 5.การผิดสัญญา", null, "28", true));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("5.1 ในกรณีต่อไปนี้ให้ถือว่าผู้กู้ผิดสัญญา ให้ผู้ให้กู้มีสิทธิบอกเลิกสัญญาได้"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("5.1.1 ผู้กู้ไม่ปฏิบัติตามสัญญาฉบับนี้ไม่ว่าข้อหนึ่งข้อใด"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("5.1.2 ผู้กู้ผิดนัดชำระคืนต้นเงินไม่ว่างวดหนึ่งงวดใดก็ตาม หรือเงินจำนวนอื่นใดที่ต้องชำระตามสัญญาฉบับนี้"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("5.1.3 ผู้กู้ให้ข้อเท็จจริง ข่าวสาร ข้อความหรือเอกสารอันเป็นเท็จ หรือปกปิด ข้อเท็จจริงซึ่งควรจะแจ้งให้ผู้ให้กู้ทราบ"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("5.1.4 ผู้กู้ไม่ปฏิบัติตามโครงการเงินทุนพลิกฟื้นวิสาหกิจขนาดย่อม ตามเอกสารแนบท้ายสัญญานี้"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("5.2 เมื่อผู้กู้ผิดสัญญาแล้วแม้ข้อหนึ่งข้อใด หรือผู้กู้ไม่ชำระหนี้ให้ถูกต้องครบถ้วนตามที่กำหนดในสัญญานี้ไม่ว่าข้อหนึ่งข้อใด หรือผิดนัดชำระหนี้งวดใด ๆให้ถือว่าเป็นการผิดนัดทั้งหมด บรรดาหนี้สินทั้งหลายที่ยังต้องชำระ\r\nอยู่ตามสัญญานี้ ไม่ว่าจะถึงกำหนดชำระแล้วหรือไม่ ให้ถือว่าเป็นอันถึงกำหนดชำระทั้งหมดทันที ผู้กู้ยินยอมให้ผู้ให้\r\nกู้คิดดอกเบี้ยจากเงินต้นที่ค้างชำระในอัตราร้อยละ 15.00 ต่อปี นับตั้งแต่วันที่ผู้กู้ตกเป็นผู้ผิดนัดตามสัญญานี้ จนกว่าจะชำระหนี้ทั้งหมดเสร็จสิ้น  พร้อมด้วยค่าเสียหายและค่าใช้จ่ายทั้งหลายอันเนื่องจากการผิดนัดชำระหนี้ของผู้กู้ รวมทั้งค่าใช้จ่าย\r\nในการเตือน เรียกร้อง ทวงถาม ดำเนินคดีและการบังคับชำระหนี้จนเต็มจำนวน\r\n"));

            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 6. การเปิดเผยข้อมูล", null, "28", true));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ในการวิเคราะห์ข้อมูลเพื่อประกอบการพิจารณาให้สินเชื่อ การแก้ไขหนี้ หรือการปรับปรุงโครงสร้างหนี้ของผู้ให้กู้แก่ผู้กู้นั้น ผู้กู้ตกลงยินยอมให้ผู้ให้กู้ตรวจสอบและใช้ข้อมูลเกี่ยวกับการเงิน ประวัติและภาระหนี้  ที่ผู้กู้มีอยู่กับสถาบันการเงิน และนิติบุคคลอื่น รวมทั้งข้อมูลเครดิตของผู้กู้ที่ได้ถูกรวบรวมไว้ที่ บริษัท ข้อมูลเครดิตแห่งชาติ จำกัด  หรือบริษัทข้อมูลเครดิตใด ๆ ตามพระราชบัญญัติการประกอบธุรกิจข้อมูลเครดิต ตลอดจนการตรวจสอบการล้มละลายและหรือ \r\nการบังคับคดีขายทอดตลาดของผู้กู้ได้ โดยไม่ต้องคำนึงว่าผู้กู้จะได้รับอนุมัติสินเชื่อ ไม่ว่าจะเป็นการให้วงเงินสินเชื่อ การแก้ไขหนี้ หรือการปรับปรุงโครงสร้างหนี้จากผู้ให้กู้หรือไม่ก็ตาม\r\n"));

            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 7. อื่นๆ", null, "28", true));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("7.1 ในระหว่างและตลอดระยะเวลาการกู้เงินตามสัญญานี้ ผู้กู้ยินยอมให้ผู้ให้กู้ หรือตัวแทนผู้ให้กู้เข้าไปตรวจสอบกิจการ ตลอดจนเอกสารหลักฐานทางบัญชีของกิจการ สรรพสมุดและเอกสารอื่นๆ ของผู้กู้ได้ตลอด"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("7.2 คู่สัญญาตกลงให้ถือเอาเอกสารที่แนบท้ายสัญญานี้ บันทึกข้อตกลง และบรรดาข้อสัญญาต่างๆ เป็นส่วนหนึ่งของสัญญานี้ที่มีผลผูกพันให้ผู้กู้จะต้องปฏิบัติตาม" +
             " ซึ่งเอกสารแนบท้ายนี้อาจจะทำเพิ่มเติมในภายหลังจากวันทำสัญญานี้ โดยให้ถือเป็นส่วนหนึ่งของสัญญานี้เช่นกัน และหากเอกสารแนบท้ายสัญญาขัดหรือแย้งกันผู้กู้ตกลงปฏิบัติตามคำวินิจฉัยของผู้ให้กู้"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("7.3 บรรดาหนังสือ จดหมาย คำบอกกล่าวใดๆ เช่น การทวงถาม การบอกเลิกสัญญา ของผู้ให้กู้ที่ส่งไปยังสถานที่ที่ระบุไว้ว่าเป็นที่อยู่ของผู้กู้ข้างต้น" +
             "หรือสถานที่อยู่ที่ผู้กู้แจ้งเปลี่ยนแปลง โดยส่งเองหรือส่งทางไปรษณีย์ลงทะเบียน หรือไม่ลงทะเบียนไม่ว่าจะมีผู้รับไว้หรือไม่มีผู้ใดยอมรับไว้" +
             "หรือส่งไม่ได้เพราะผู้กู้ย้ายสถานที่อยู่ไปโดยมิได้แจ้งให้ผู้ให้กู้ทราบให้ไว้นั้นหาไม่พบ หรือถูกรื้อถอนทำลายทุกๆ กรณีดังกล่าวให้ถือว่าผู้กู้ได้รับโดยชอบแล้ว"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("7.4 การสละสิทธิ์ตามสัญญานี้ ในคราวหนึ่งคราวใดของผู้ให้กู้ หรือการที่ผู้ให้กู้มิได้ใช้สิทธิ์ที่มีอยู่ ไม่ถือเป็นการสละสิทธิ์ของผู้ให้กู้ในคราวต่อไปและไม่มีผลกระทบต่อการใช้สิทธิของผู้ให้กู้ในคราวต่อไป"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("7.5 หากข้อกำหนด และ/หรือเงื่อนไขข้อใดข้อหนึ่งของสัญญานี้ตกเป็นโมฆะ หรือใช้บังคับไม่ได้ตามกฎหมาย ให้ข้อกำหนดและเงื่อนไขอื่น ๆ ยังคงมีผลใช้บังคับได้ต่อไปได้ โดยแยกต่างหากจากส่วนที่เป็นโมฆะหรือไม่สมบูรณ์นั้น"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ผู้กู้ได้ตรวจ อ่าน และเข้าใจข้อความในสัญญานี้โดยละเอียดโดยตลอดแล้ว เห็นว่าถูกต้องตามเจตนาทุกประการ จึงได้ลงลายมือชื่อพร้อมประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยาน ณ วัน เดือน ปี ที่ระบุไว้ข้างต้น"));

            body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ........................................................................ผู้กู้"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("(................................................................................)"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ.......................................................................คู่สมรสให้ความยินยอม"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("(................................................................................)"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ.......................................................................คู่สมรสให้ความยินยอม"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("(...............................................................................)"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ......................................................................พยาน"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("(...............................................................................)"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ.....................................................................พยาน"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("(...............................................................................)"));

            // next page
            body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

            body.AppendChild(WordServiceSetting.CenteredParagraph("คำรับรองสถานภาพการสมรส"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้าพเจ้า………………………………………………………………………………………………………………………………………….\r\nขอรับรองว่าสถานภาพการสมรสของข้าพเจ้าปัจจุบันมีสถานะ\r\n"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("“ข้าพเจ้าขอรับรองว่าสถานภาพการสมรสที่แจ้งในหนังสือฉบับนี้เป็นความจริงทุกประการหากไม่เป็นความจริงแล้ว ความเสียหายใด ๆ ที่จะเกิดกับผู้ให้การอุดหนุน ข้าพเจ้ายินยอมรับผิดชดใช้ให้แก่ผู้ให้การอุดหนุนทั้งสิ้น”"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ.............................................................รับรอง"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("(............................................................)"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ....................................................พยาน          ลงชื่อ ........................................................พยาน"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("(............................................................)                 (.........................................................)"));

            body.AppendChild(WordServiceSetting.EmptyParagraph());
            body.AppendChild(WordServiceSetting.RightParagraph("........................................................./ผู้พิมพ์"));
            body.AppendChild(WordServiceSetting.RightParagraph("........................................................./ผู้ตรวจ"));


            // --- Add header for first page (empty) ---
            WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);
        }
        stream.Position = 0;
        return stream.ToArray();
    }
    #endregion   สสว. สัญญาเงินกู้ยืม โครงการพลิกฟื้นวิสาห

}
