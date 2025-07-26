using BatchAndReport.Entities;
using BatchAndReport.Models;
using BatchAndReport.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

public class WordEContract_AllowanceService 
{
    private readonly WordServiceSetting _w;

    public WordEContract_AllowanceService(WordServiceSetting ws)
    {
        _w = ws;
    }
    #region  สสว. สัญญารับเงินอุดหนุน
    public byte[] OnGetWordContact_Allowance()
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
            body.AppendChild(WordServiceSetting.RightParagraph("เลขที่สัญญา ............................"));
            body.AppendChild(WordServiceSetting.EmptyParagraph());
            body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("สัญญารับเงินอุดหนุน", "FF0000")); // Blue
            body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("ตามแนวทางการดำเนินโครงการวิสาหกิจขนาดกลางและขนาดย่อมต่อเนื่อง", "FF0000")); // Red
            body.AppendChild(WordServiceSetting.CenteredBoldParagraph("ที่ศูนย์ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม"));
            body.AppendChild(WordServiceSetting.CenteredBoldParagraph("วันที่...................................................."));

            // 3. Fillable lines (using underlines)
            //  body.AppendChild(EmptyParagraph());
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้าพเจ้า ...................................................................................................................."));
            body.AppendChild(WordServiceSetting.JustifiedParagraph("อายุ ......... ปี สัญชาติ .................. สำนักงาน/บ้านตั้งอยู่เลขที่.................. อาคาร..........................................."));
            body.AppendChild(WordServiceSetting.JustifiedParagraph("หมู่ที่...........ตรอก/ซอย..........................ถนน...........................ตำบล/แขวง.................. ..................."));
            body.AppendChild(WordServiceSetting.JustifiedParagraph("เขต/อำเภอ................... จังหวัด.................ทะเบียนนิติบุคคลเลขที่/เลขประจำตัวประชาชนที่............................................"));
            body.AppendChild(WordServiceSetting.JustifiedParagraph("จดทะเบียนเป็นนิติบุคคลเมื่อวันที่ .........................................."));

            // 4. Main body (sample)
            //   body.AppendChild(EmptyParagraph());
            body.AppendChild(WordServiceSetting.NormalParagraph("ซึ่งต่อไปนี้จะเรียกบุคคลผู้มีนามตามที่ระบุข้างต้นทั้งหมดว่า \"ผู้รับการอุดหนุน\" ได้ทำสัญญาฉบับนี้ให้ไว้แก่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งต่อไปนี้จะเรียกว่า \"ผู้ให้การอุดหนุน\" โดยมีสาระสำคัญดังนี้"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ1. ผู้รับการอุดหนุนได้ขอรับความช่วยเหลือผ่านการอุดหนุนตามมาตรการฟื้นฟูกิจการวิสาหกิจ ขนาดกลางและขนาดย่อมจากผู้ให้การอุดหนุนเป็นจำนวนเงิน ..................... บาท (...........................) ปลอดการชำระเงินต้น ................. เดือน โดยไม่มีดอกเบี้ย แต่มีภาระต้องชำระคืนเงินต้น "));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ2. ผู้ให้การอุดหนุนจะให้ความช่วยเหลือด้วยการให้เงินอุดหนุนแก่ผู้รับการอุดหนุน ด้วยการนำเงินหรือโอนเงินเข้าบัญชีธนาคารกรุงไทย จำกัด (มหาชน) สาขา ..................................... เลขที่บัญชี ................................................. ชื่อบัญชี .......................... ซึ่งเป็นบัญชีของผู้รับการอุดหนุน จำนวนเงิน ........................... บาท (..............................) และให้ถือว่าผู้รับการอุดหนุนได้รับเงินอุดหนุนตามสัญญานี้ไปจากผู้ให้การอุดหนุนแล้ว ในวันที่เงินเข้าบัญชีของผู้รับการอุดหนุนดังกล่าว"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ3. ห้ามผู้รับการอุดหนุนนำเงินอุดหนุนไปชำระหนี้เดิมที่มีอยู่ก่อนทำสัญญานี้"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ4. ผู้รับการอุดหนุนยินยอมให้ธนาคารกรุงไทย จำกัด (มหาชน) ซึ่งกระทำการแทนผู้ให้การอุดหนุน หักเงินอุดหนุนที่จะได้จากผู้ให้การอุดหนุนเป็นค่าใช้จ่ายหรือค่าธรรมเนียมในการโอนเงินเข้าบัญชีของผู้รับการอุดหนุน ซึ่งธนาคารกรุงไทย จำกัด (มหาชน) เรียกเก็บตามระเบียบของธนาคารได้ โดยไม่ต้องบอกกล่าวหรือแจ้งให้ผู้รับการอุดหนุนทราบล่วงหน้า และให้ถือว่าผู้รับการอุดหนุนได้รับเงินตามจำนวนที่เบิกไปครบถ้วนแล้ว"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ5. ผู้รับการอุดหนุนตกลงผ่อนชำระเงินต้นคืนให้แก่ผู้ให้การอุดหนุนเป็นรายเดือน (งวด) ๆ ละ ไม่น้อยกว่า ....................... บาท (.....................................) ด้วยการโอนเข้าบัญชีตามที่ระบุไว้ในข้อ 2 โดยชำระเงินต้นงวดแรกในเดือนที่ ....................... นับถัดจากวันที่ได้รับเงินอุดหนุน และงวดถัดไปทุกวันที่ .................. ของเดือนจนกว่าจะชำระเสร็จสิ้น  แต่ทั้งนี้จะต้องชำระให้เสร็จสิ้นไม่เกินกว่า .............. ปี (...........) นับแต่วันที่ได้รับเงินอุดหนุน"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ6. การชำระเงินคืนตาม"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ5. ผู้รับการอุดหนุนตกลงจะนำเงินเข้าบัญชีเงินฝากของผู้รับการอุดหนุน ที่เปิดบัญชีไว้กับธนาคารกรุงไทย จำกัด (มหาชน) ตามข้อ 2 โดยผู้รับการอุดหนุนยินยอมให้ ธนาคารกรุงไทย จำกัด (มหาชน) ซึ่งดำเนินการแทนผู้ให้การอุดหนุน หักเงินจากบัญชีของผู้รับการอุดหนุนดังกล่าวเพื่อชำระคืนเงินอุดหนุนแก่ผู้ให้การอุดหนุน"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ6. การชำระเงินคืนตามข้อ 5 ผู้รับการอุดหนุนตกลงจะนำเงินเข้าบัญชีเงินฝากของผู้รับการอุดหนุนที่เปิดบัญชีไว้กับธนาคารกรุงไทย จำกัด (มหาชน) ตามข้อ 2 โดยผู้รับการอุดหนุนยินยอมให้ ธนาคารกรุงไทย จำกัด (มหาชน) ซึ่งดำเนินการแทนผู้ให้การอุดหนุน หักเงินจากบัญชีของผู้รับการอุดหนุนดังกล่าวเพื่อชำระคืนเงินอุดหนุนแก่ผู้ให้การอุดหนุน ในแต่ละงวดเดือน พร้อมทำการโอนเงินที่พักจากบัญชีของผู้รับการอุดหนุนมอบเข้าบัญชีของผู้ให้การอุดหนุนที่เปิดบัญชีไว้กับธนาคารกรุงไทย จำกัด (มหาชน) สาขา .....องค์การตลาดเพื่อเกษตรกร (จตุจักร)..... บัญชีออมทรัพย์ ชื่อบัญชีสำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เลขที่บัญชี ....035-1-52709-5.....เพื่อชำระหนี้คืนเงินอุดหนุนแก่ผู้ให้การอุดหนุนตามข้อตกลงในแต่ละงวดเดือน"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ไม่ว่าผู้รับการอุดหนุนจะได้จัดทำหนังสือยินยอมให้ธนาคารกรุงไทย จำกัด (มหาชน) หักบัญชีเงินฝาก ตามวรรคหนึ่งหรือไม่ก็ตาม โดยสัญญนี้ผู้รับการอุดหนุนให้ถือว่าเป็นการทำหนังสือยินยอมให้ธนาคารกรุงไทย จำกัด (มหาชน) หักบัญชีเงินฝากตามวรรคหนึ่งด้วย"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ7. ผู้รับการอุดหนุนตกลงยินยอมให้ธนาคารกรุงไทย จจำกัด (มหาชน) ซึ่งกระทำการแทนผู้ให้การอุดหนุนหักเงินที่ผู้รับการอุดหนุนได้โอนเข้าบัญชีตามข้อ 2 เพื่อชำระคืนเป็นค่าใช้จ่ายหรือค่าธรรมเนียมในการโอนเงิน"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ8. ในระหว่างและตลอดระยะเวลาการตามสัญญาฉบับนี้ ผู้รับการอุดหนุนจะต้องรายงานผลการประกอบกิจการมายังผู้ให้การอุดหนุนหรือศูนย์ให้บริหร SMEs ครบวงจร ในจังหวัดที่ผู้รับการอุดหนุนมีภูมิลำเนาอยู่หรือพื้นที่ใกล้เคียงหรือหน่วยงานอื่นใดที่ผู้ให้การอุดหนุนมอบหมาย ตามหลักเกณฑ์และวิธีการที่ผู้ให้การอุดหนุนกำหนด ไม่น้อยกว่าเดือนละหนึ่งครั้ง"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ9. กรณีต่อไปนี้ให้ถือว่าผู้รับการอุดหนุนปฏิบัติผิดสัญญา"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("9.1 ผู้รับการอุดหนุนผิดนัดชำระคืนเงินอุดหนุนไม่ว่างวดหนึ่งวดใดก็ตาม หรือไม่ชำระคืนเงินอุดหนุนภายในกำหนดระยะเวลาที่กำหนดในสัญญานี้ หรือเงินจำนวนอื่นใดที่ต้องชำระตามสัญญาฉบับนี้"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("9.2 ผู้รับการอุดหนุนใช้เงินอุดหนุนผิดไปจากเงื่อนไขตามสัญญา หรือผิดสัญญาแม้ข้อใดข้อหนึ่ง หรือไม่รายงานการดำเนินธุรกิจให้ผู้ให้การอุดหนุนทราบตามข้อ 8 หรือตรวจสอบในภายหลังแล้วพบว่ามีการแจ้งคุณสมบัติ หรือส่งเอกสารเป็นเท็จแก่ผู้ให้การอุดหนุน ๆ มีสิทธิบอกเลิกสัญญาได้"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ10. อื่นๆ"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("10.1 ในระหว่างและตลอดระยะเวลาตามสัญญานี้ ผู้รับการอุดหนุนยินยอมให้ผู้ให้การอุดหนุน หรือตัวแทนผู้ให้การอุดหนุนเข้าไปตรวจสอบติดตามการดำเนินธุรกิจ ตลอดจนเอกสารหลักฐานทางบัญชีของกิจการ สรรพเอกสารอื่น ๆ ของผู้รับการอุดหนุนได้ตลอด"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("10.2 คู่สัญญาตกลงให้ถือเอาเอกสารที่แนบท้ายสัญญานี้ บันทึกข้อตกลง และบรรดาข้อสัญญาต่าง ๆ  เป็นส่วนหนึ่งของสัญญานี้ที่มีผลผูกพันให้ผู้รับการอุดหนุนจะต้องปฏิบัติตาม ซึ่งเอกสารแนบท้ายนี้อาจจะทำเพิ่มเติมในภายหลังจากวันทำสัญญานี้ โดยให้ถือเป็นส่วนหนึ่งของสัญญานี้เช่นกัน และหากเอกสารแนบท้ายสัญญาขัดหรือแย้งกันผู้รับการอุดหนุนตกลงปฏิบัติตามคำวินิจฉัยของผู้ให้การอุดหนุน"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("10.3 บรรดาหนังสือ จดหมาย คำบอกกล่าวใด ๆ เช่น การทวงถาม การบอกเลิกสัญญา ของผู้ให้การอุดหนุนหรือผู้ที่ได้รับมอบหมายส่งไปยังสถานที่ที่ระบุไว้เป็นที่อยู่ของผู้รับการอุดหนุนข้างต้น หรือสถานที่อยู่ที่ผู้รับการอุดหนุนแจ้งเปลี่ยนแปลง โดยส่งเองหรือส่งทางไปรษณีย์ลงทะเบียน หรือไม่ลงทะเบียน ไม่ว่าจะมีผู้รับไว้ หรือไม่มีผู้ใดยอมรับไว้ หรือส่งไม่ได้เพราะผู้รับการอุดหนุนย้ายสถานที่อยู่ไปโดยมิได้แจ้งให้ผู้ให้การอุดหนุนทราบหรือหาไม่พบ หรือถูกรื้อถอนทำลายทุก ๆ กรณีดังกล่าวให้ถือว่าผู้รับการอุดหนุนได้รับโดยชอบแล้ว"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("10.4 การสละสิทธิ์ตามสัญญานี้ ในคราวหนึ่งคราวใดของผู้ให้การอุดหนุน หรือการที่ผู้ให้การอุดหนุนมิได้ ใช้สิทธิ์ที่มีอยู่ ไม่ถือเป็นการสละสิทธิ์ของผู้ให้การอุดหนุนในคราวต่อไปและไม่มีผลกระทบต่อการใช้สิทธิของผู้ให้การอุดหนุน ในคราวต่อไป"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("10.5 หากข้อกำหนด และ/หรือเงื่อนไขข้อใดข้อหนึ่งของสัญญานี้ตกเป็นโมฆะ หรือใช้บังคับไม่ได้ตามกฎหมาย ให้ข้อกำหนดและเงื่อนไขอื่น ๆ ยังคงมีผลใช้บังคับได้ต่อไปได้ โดยแยกต่างหากจากส่วนที่เป็นโมฆะหรือไม่สมบูรณ์นั้น"));
            body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สัญญานี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน คู่สัญญาทั้งสองฝ่ายได้ตรวจ อ่าน และเข้าใจข้อความในสัญญานี้โดยละเอียดแล้ว เห็นว่าถูกต้องตามเจตนาทุกประการ จึงได้ลงลายมือชื่อพร้อมประทับตรา (ถ้ามี) ไว้เป็นสำคัญ ต่อหน้าพยาน ณ วัน เดือน ปี ที่ระบุไว้ข้างต้น"));

            body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ........................................................................ผู้ให้การอุดหนุน"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("(................................................................................)"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ.......................................................................ผู้รับการอุดหนุน"));
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
            body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ....................................................พยาน                    ลงชื่อ ........................................................พยาน"));
            body.AppendChild(WordServiceSetting.CenteredParagraph("(............................................................)                                 (.........................................................)"));

            body.AppendChild(WordServiceSetting.EmptyParagraph());
            body.AppendChild(WordServiceSetting.RightParagraph("........................................................./ผู้พิมพ์"));
            body.AppendChild(WordServiceSetting.RightParagraph("........................................................./ผู้ตรวจ"));


            WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);
        }
        stream.Position = 0;
        return stream.ToArray();
    }
    #endregion  สสว. สัญญารับเงินอุดหนุน
}
