﻿using BatchAndReport.DAO;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading.Tasks;

public class WordEContract_PersernalProcessService
{
    private readonly WordServiceSetting _w;
    private readonly E_ContractReportDAO _eCon;

    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    public WordEContract_PersernalProcessService(WordServiceSetting ws
            , E_ContractReportDAO eContractReportDAO
          
      , IConverter pdfConverter
        )
    {
        _w = ws;
        _eCon = eContractReportDAO;
        _pdfConverter = pdfConverter; // กำหนดค่า PDF Converter

    }
    #region 4.1.1.2.4.บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล
    public async Task<byte[]> OnGetWordContact_PersernalProcessService(string id)
    {
        var result = await _eCon.GetPDPAAsync(id);

        if (result == null)
        {
            throw new Exception("PDPA data not found.");
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
                // Add this line for the header logo path:
                var imageHeaderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo1.jpg");

                // Add image part and feed image data
                var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg, "rIdLogo");
                using (var imgStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read, FileShare.Read))
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
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("โครงการ.."+result.Project_Name+"...", "44"));

                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("ระหว่าง", "32"));

                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม กับ "+result.Contract_Organization, "32"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("---------------------------------", "36"));

                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // --- 3. Main contract body ---
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อตกลงการประมวลผลข้อมูลส่วนบุคคล (“ข้อตกลง”) ฉบับนี้ทำขึ้น เมื่อวันที่..."+result.Master_Contract_Sign_Date??""+"... ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("โดยที่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สสว.” ฝ่ายหนึ่ง " +
                    "ได้ตกลงใน....."+result.Master_Contract_Number??""+"... " +
                    "สัญญาเลขที่ ...."+result.Contract_Number??""+".... " +
                    "ฉบับลงวันที่ ..."+result.Master_Contract_Sign_Date??""+"... " +
                    "ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “(บันทึกความร่วมมือ/สัญญา)” กับ "+result.Contract_Organization +
                    "ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “" + result.Contract_Organization + "” อีกฝ่ายหนึ่ง", null, "32"));




                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ตามที่ (ระบุชื่อบันทึกความร่วมมือ/สัญญาหลัก) ดังกล่าวกำหนดให้ สสว."+result.OSMEP_ScopeRightsDuties+" ซึ่งในการดำเนินการดังกล่าวประกอบด้วยการมอบหมายหรือแต่งตั้งให้" + result.Contract_Organization + "เป็นผู้ดำเนินการกระบวนการเก็บรวบรวม ใช้ หรือเปิดเผย (“ประมวลผล”) ข้อมูลส่วนบุคคลแทนหรือในนามของ สสว. ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สสว. ในฐานะผู้ควบคุมข้อมูลส่วนบุคคลเป็นผู้มีอำนาจตัดสินใจ กำหนดรูปแบบและกำหนดวัตถุประสงค์ในการประมวลผลข้อมูลส่วนบุคคล ได้.....(มอบหมาย/แต่งตั้ง/จ้าง/อื่น ๆ).....ให้ " + result.Contract_Organization + " ในฐานะผู้ประมวลผลข้อมูลส่วนบุคคล ดำเนินการเพื่อวัตถุประสงค์ดังต่อไปนี้", null, "32"));


                var conPurpose = await _eCon.GetPDPA_ObjectivesAsync(id);
                if (conPurpose != null && conPurpose.Count > 0)
                {
                    int rowp = 1;
                    foreach (var purpose in conPurpose)
                    {
                        body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs(purpose.Objective_Description, null, "32"));
                        rowp++;
                    }
                }

                var conAgreement = await _eCon.GetPDPA_AgreementListAsync(id);
                if (conAgreement != null && conAgreement.Count > 0)
                {
                    int rowa = 1;
                    foreach (var agreement in conAgreement)
                    {
                        body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs( agreement.PD_Detail, null, "32"));
                        rowa++;
                    }
                }

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ด้วยเหตุนี้ ทั้งสองฝ่ายจึงตกลงจัดทำข้อตกลงฉบับนี้ และให้ถือข้อตกลงฉบับนี้เป็นส่วนหนึ่งของ...."+result.Master_Contract_Number??""+"...เพื่อเป็นหลักฐานการควบคุมดูแลการประมวลผลข้อมูลส่วนบุคคลที่ สสว. มอบหมายหรือแต่งตั้งให้ " + result.Contract_Organization + " ดำเนินการ อันเนื่องมาจากการดำเนินการตามหน้าที่และความรับผิดชอบตาม..."+result.Master_Contract_Number??""+"..." +
                    "ฉบับลงวันที่ ."+result.Master_Contract_Sign_Date??""+". และเพื่อดำเนินการให้เป็นไปตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒ และกฎหมายอื่นๆ ที่ออกตามความในพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒ ซึ่งต่อไปในข้อตกลงฉบับนี้ รวมเรียกว่า “กฎหมายคุ้มครองข้อมูลส่วนบุคคล” ทั้งที่มีผลใช้บังคับอยู่ ณ วันทำข้อตกลงฉบับนี้และที่จะมีการเพิ่มเติมหรือแก้ไขเปลี่ยนแปลงในภายหลัง โดยมีรายละเอียดดังนี้ ", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("1. " + result.Contract_Organization + " รับทราบว่า ข้อมูลส่วนบุคคล หมายถึง ข้อมูลเกี่ยวกับบุคคลธรรมดาซึ่งทำให้สามารถระบุตัวบุคคลนั้นได้ไม่ว่าทางตรงหรือทางอ้อม โดย" + result.Contract_Organization + "จะดำเนินการ ตามที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลกำหนด เพื่อคุ้มครองให้การประมวลผลข้อมูลส่วนบุคคลเป็นไปอย่างเหมาะสมและถูกต้องตามกฎหมาย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("โดยในการดำเนินการตามข้อตกลงนี้  " + result.Contract_Organization + " .จะประมวลผลข้อมูลส่วนบุคคลเมื่อได้รับคำสั่งที่เป็นลายลักษณ์อักษรจาก สสว. แล้วเท่านั้น ทั้งนี้ เพื่อให้ปราศจากข้อสงสัย การดำเนินการประมวลผลข้อมูลส่วนบุคคลโดย" + result.Contract_Organization + "ตามหน้าที่และความรับผิดชอบตาม.."+result.Master_Contract_Number??""+"..ถือเป็นการได้รับคำสั่งที่เป็นลายลักษณ์อักษรจาก สสว. แล้ว", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("2. " + result.Contract_Organization + "จะกำหนดให้การเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้ถูกจำกัดเฉพาะเจ้าหน้าที่ และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ได้รับมอบหมาย มีหน้าที่เกี่ยวข้องหรือมีความจำเป็นในการเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้เท่านั้น และจะดำเนินการเพื่อให้พนักงาน และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ได้รับมอบหมายจาก " + result.Contract_Organization + " ทำการประมวลผลและรักษาความลับของข้อมูลส่วนบุคคลด้วยมาตรฐานเดียวกัน", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3. " + result.Contract_Organization + "จะควบคุมดูแลให้เจ้าหน้าที่ และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ปฏิบัติหน้าที่ในการประมวลผลข้อมูลส่วนบุคคล ปฏิบัติตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลอย่างเคร่งครัด และดำเนินการประมวลผลข้อมูลส่วนบุคคลตามวัตถุประสงค์ของการดำเนินการตามข้อตกลงฉบับนี้เท่านั้น โดยจะไม่ทำซ้ำ คัดลอก ทำสำเนา บันทึกภาพข้อมูลส่วนบุคคลไม่ว่าทั้งหมดหรือแต่บางส่วนเป็นอันขาด เว้นแต่เป็นไปตามเงื่อนไขของบันทึกความร่วมมือหรือสัญญา หรือกฎหมายที่เกี่ยวข้องจะระบุหรือบัญญัติไว้เป็นประการอื่น", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("4. " + result.Contract_Organization + "จะดำเนินการเพื่อช่วยเหลือหรือสนับสนุน สสว. ในการตอบสนองต่อคำร้องที่เจ้าของข้อมูลส่วนบุคคลแจ้งต่อ สสว. อันเป็นการใช้สิทธิของเจ้าของข้อมูลส่วนบุคคลตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลในส่วนที่เกี่ยวข้องกับการประมวลผลข้อมูลส่วนบุคคลในขอบเขตของข้อตกลงฉบับนี้ ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("อย่างไรก็ดี ในกรณีที่เจ้าของข้อมูลส่วนบุคคลยื่นคำร้องขอใช้สิทธิดังกล่าวต่อ " + result.Contract_Organization + "โดยตรง", null, "32"));
                body.AppendChild(WordServiceSetting.JustifiedParagraph(result.Contract_Organization + "..จะดำเนินการแจ้งและส่งคำร้องดังกล่าวให้แก่ สสว. ทันที โดย", "32"));
                body.AppendChild(WordServiceSetting.JustifiedParagraph(result.Contract_Organization + "..จะไม่เป็นผู้ตอบสนองต่อคำร้องดังกล่าว เว้นแต่ สสว. จะได้มอบหมายให้", "32"));
                body.AppendChild(WordServiceSetting.JustifiedParagraph(result.Contract_Organization + "..ดำเนินการเฉพาะเรื่องที่เกี่ยวข้องกับคำร้องดังกล่าว", "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("5. " + result.Contract_Organization + "จะจัดทำและเก็บรักษาบันทึกรายการของกิจกรรมการประมวลผลข้อมูลส่วนบุคคล (Record of Processing Activities) ทั้งหมดที่  " + result.Contract_Organization + "  ประมวลผลในขอบเขตของข้อตกลงฉบับนี้ และจะดำเนินการส่งมอบบันทึกรายการดังกล่าวให้แก่ สสว. ทุก.....(ระบุความถี่ของการส่งมอบบันทึกรายการ เช่น ทุกสัปดาห์หรือทุกเดือน).... และ/หรือทันทีที่ สสว. ร้องขอ", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("6. " + result.Contract_Organization + "จะจัดให้มีและคงไว้ซึ่งมาตรการรักษาความปลอดภัยสำหรับการประมวลผลข้อมูลที่มีความเหมาะสมทั้งในเชิงองค์กรและเชิงเทคนิคตามที่คณะกรรมการคุ้มครองข้อมูลส่วนบุคคลได้ประกาศกำหนดและ/หรือตามมาตรฐานสากล โดยคำนึงถึงลักษณะ ขอบเขต และวัตถุประสงค์ของการประมวลผลข้อมูลตามที่กำหนดในข้อตกลงฉบับนี้เป็นสำคัญ เพื่อคุ้มครองข้อมูลส่วนบุคคลจากความเสี่ยงอันเกี่ยวเนื่องกับการประมวลผลข้อมูลส่วนบุคคล เช่น ความเสียหายอันเกิดจากการละเมิด อุบัติเหตุ การลบ ทำลาย สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้ เปิดเผยหรือโอนข้อมูลส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย เป็นต้น", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("7. เว้นแต่กฎหมายที่เกี่ยวข้องจะบัญญัติไว้เป็นประการอื่น " + result.Contract_Organization + " จะทำการลบหรือทำลายข้อมูลส่วนบุคคลที่ทำการประมวลผลภายใต้ข้อตกลงฉบับนี้ภายใน."+result.RetentionPeriodDays+".วัน นับแต่วันที่ดำเนินการประมวลผลเสร็จสิ้น หรือวันที่ สสว. และ  " + result.Contract_Organization + "ได้ตกลงเป็นลายลักษณ์อักษรให้ยกเลิก..."+result.Master_Contract_Number??""+"...แล้วแต่กรณีใดจะเกิดขึ้นก่อน ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("นอกจากนี้ ในกรณีปรากฏว่า " + result.Contract_Organization + " หมดความจำเป็นจะต้องเก็บรักษาข้อมูลส่วนบุคคลตามข้อตกลงฉบับนี้ก่อนสิ้นระยะเวลาตามวรรคหนึ่ง " + result.Contract_Organization + " จะทำการลบหรือทำลายข้อมูลส่วนบุคคลตามข้อตกลงฉบับนี้ทันที", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("8. กรณีที่ " + result.Contract_Organization + " พบพฤติการณ์ใด ๆ ที่มีลักษณะที่กระทบต่อการรักษาความปลอดภัยของข้อมูลบุคคลที่ " + result.Contract_Organization + " ..ประมวลผลภายใต้ข้อตกลงฉบับนี้ ซึ่งอาจก่อให้เกิดความเสียหายจากการละเมิด อุบัติเหตุ การลบ ทำลาย สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้ เปิดเผยหรือโอนข้อมูลส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย แล้ว " + result.Contract_Organization + " จะดำเนินการแจ้งให้ สสว. ทราบโดยทันทีภายในเวลาไม่เกิน....(ระบุเวลาเป็นหน่วยชั่วโมงที่คู่สัญญาต้องแจ้งเหตุแก่ สสว. เช่น ภายใน 24 ชั่วโมงหรือ 48 ชั่วโมง ทั้งนี้ไม่ควรเกิน 48 ชั่วโมงเนื่องจาก สสว. ในฐานะผู้ควบคุมข้อมูลส่วนบุคคลมีหน้าที่ต้องแจ้งเหตุดังกล่าวแก่คณะกรรมการคุ้มครองข้อมูลส่วนบุคคลภายใน 72 ชั่วโมง).... ชั่วโมง", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("9. การแจ้งถึงเหตุการละเมิดข้อมูลส่วนบุคคลที่เกิดขึ้นภายใต้ข้อตกลงนี้ " + result.Contract_Organization + " จะใช้มาตรการตามที่เห็นสมควรในการระบุถึงสาเหตุของการละเมิด และป้องกันปัญหาดังกล่าวมิให้เกิดซ้ำ และจะให้ข้อมูลแก่ สสว. ภายใต้ขอบเขตที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลได้กำหนด ดังต่อไปนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("- รายละเอียดของลักษณะและผลกระทบที่อาจเกิดขึ้นของการละเมิด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("- มาตรการที่ถูกใช้เพื่อลดผลกระทบของการละเมิด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("- ข้อมูลอื่น ๆ เกี่ยวข้องกับการละเมิด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("10. หน้าที่และความรับผิดของ " + result.Contract_Organization + " ในการปฏิบัติตามข้อตกลงจะสิ้นสุดลงนับแต่วันที่ปฏิบัติงานที่ตกลงเสร็จสิ้น หรือ วันที่ " + result.End_Date + " และ สสว. ได้ตกลงเป็นลายลักษณ์อักษรให้ยกเลิก..."+result.Master_Contract_Number??""+"...แล้วแต่กรณีใดจะเกิดขึ้นก่อน อย่างไรก็ดี การสิ้นผลลงของข้อตกลงนี้ ไม่กระทบต่อหน้าที่ของ " + result.Contract_Organization + " ในการลบหรือทำลายข้อมูลส่วนบุคคลตามที่ได้กำหนดในข้อ 7 ของข้อตกลงฉบับนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ทั้งสองฝ่ายได้อ่านและเข้าใจข้อความโดยละเอียดตลอดแล้ว เพื่อเป็นหลักฐานแห่งการนี้ ทั้งสองฝ่ายจึงได้ลงนามไว้เป็นหลักฐานต่อหน้าพยาน ณ วัน เดือน ปี ที่ระบุข้างต้น", null, "32"));



                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // --- 6. Signature lines ---
                // --- 6. Signature lines ---
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // Main signature table: two columns
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
                    // First row: signatures
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.RightParagraph("ลงชื่อ.................................................................")
                        ),
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.RightParagraph("ลงชื่อ.................................................................")
                        )
                    ),
                    // Second row: names
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.CenteredParagraph("(............................................................)")
                        ),
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.CenteredParagraph("(...........................ชื่อคู่สัญญา............................)")
                        )
                    ),
                    // Third row: organization/role
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.CenteredBoldColoredParagraph("ผู้ควบคุมข้อมูลสำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม หรือผู้ที่ผู้ควบคุมมอบหมาย", "#00000")
                        ),
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.CenteredParagraph("............................................................")
                        )
                    )
                );
                body.AppendChild(signatureTable);
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // Witness table: two columns
                var witnessTable = new Table(
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
                    // First row: witness signatures
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.RightParagraph("ลงชื่อ.................................................................พยาน")
                        ),
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.RightParagraph("ลงชื่อ.................................................................พยาน")
                        )
                    ),
                    // Second row: witness names
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.CenteredParagraph("(............................................................)")
                        ),
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.CenteredParagraph("(...........................ชื่อคู่สัญญา............................)")
                        )
                    ),
                    // Third row: organization/role
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.CenteredBoldColoredParagraph("(สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม)", "#000000")
                        ),
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.CenteredBoldColoredParagraph("(ชื่อคู่สัญญา)", "#000000")
                        )
                    )
                );
                body.AppendChild(witnessTable);
                body.AppendChild(WordServiceSetting.EmptyParagraph());



                body.AppendChild(WordServiceSetting.EmptyParagraph());


                // --- 7. Add header/footer if needed ---
                // Add image to header part with correct relationship type
                WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);
            
            }
            stream.Position = 0;
            return stream.ToArray();
        }
      
    }

    public async Task<byte[]> OnGetWordContact_PersernalProcessService_HtmlToPDF(string id)
    {
        try {
         
            var result = await _eCon.GetPDPAAsync(id);
            var conPurpose = await _eCon.GetPDPA_ObjectivesAsync(id);
            var conAgreement = await _eCon.GetPDPA_AgreementListAsync(id);

            string strEndate = "";

            //     string strEndate = CommonDAO.ToThaiDateStringCovert(result.End_Date??DateTime.) ?? "";

            if (result == null)
            {
                throw new Exception("PDPA data not found.");
            }

            // Logo
            var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
            string logoBase64 = "";
            if (System.IO.File.Exists(logoPath))
            {
                var bytes = System.IO.File.ReadAllBytes(logoPath);
                logoBase64 = Convert.ToBase64String(bytes);
            }

            // Font
            var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");

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
        body {{
            font-size: 22px;
            font-family: 'THSarabunNew', Arial, sans-serif;
        }}
        .t-16 {{
            font-size: 2.0em;
        }}
        .t-18 {{
            font-size: 2.5em;
        }}
        .t-22 {{
            font-size: 3.0em;
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
        .sign-double {{ display: flex; }}
        .text-center-right-brake {{
            margin-left: 50%;
            word-break: break-all;
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
    </style>
</head>
<body>
    <table class='logo-table'>
        <tr>
            <td style='width:60%; text-align:left; vertical-align:top;'>
                <img src='data:image/jpeg;base64,{logoBase64}' width='240' height='80' />
            </td>
            <td style='width:40%; text-align:center; vertical-align:top;'>
                <!-- Contract code box or second logo if needed -->
            </td>
        </tr>
    </table>
    <div class='t-22 text-center'><b>ข้ออตกลงการประมวลผลข้อมูลส่วนบุคคล</b></div>
    <div class='t-22 text-center'><b>(Data Processing Agreement)</b></div>
    <div class='t-22 text-center'><b>โครงการ {result.Project_Name ?? ""}</b></div>
    <div class='t-18 text-center'><b>ระหว่าง</div>
    <div class='t-18 text-center'><b>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม กับ {result.Contract_Organization ?? ""}</b></div>
    <div class='t-18 text-center'>---------------------------------</div>
    <div class='t-16 tab3'>
        ข้อตกลงการประมวลผลข้อมูลส่วนบุคคล (“ข้อตกลง”) ฉบับนี้ทำขึ้น เมื่อวันที่ {result.Master_Contract_Sign_Date?.ToString("dd/MM/yyyy") ?? ""} ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม
    </div>
    <div class='t-16 tab3'>
        โดยที่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สสว.” ฝ่ายหนึ่ง ได้ตกลงใน {result.Master_Contract_Number ?? ""} สัญญาเลขที่ {result.Contract_Number ?? ""} ฉบับลงวันที่ {result.Master_Contract_Sign_Date?.ToString("dd/MM/yyyy") ?? ""} ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “(บันทึกความร่วมมือ/สัญญา)” กับ {result.Contract_Organization ?? ""} ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “{result.Contract_Organization ?? ""}” อีกฝ่ายหนึ่ง
    </div>
    <div class='t-16 tab3'>
        ตามที่ (ระบุชื่อบันทึกความร่วมมือ/สัญญาหลัก) ดังกล่าวกำหนดให้ สสว.{result.OSMEP_ScopeRightsDuties ?? ""} ซึ่งในการดำเนินการดังกล่าวประกอบด้วยการมอบหมายหรือแต่งตั้งให้ {result.Contract_Organization ?? ""} เป็นผู้ดำเนินการกระบวนการเก็บรวบรวม ใช้ หรือเปิดเผย (“ประมวลผล”) ข้อมูลส่วนบุคคลแทนหรือในนามของ สสว.
    </div>
    <div class='t-16 tab3'>
        สสว. ในฐานะผู้ควบคุมข้อมูลส่วนบุคคลเป็นผู้มีอำนาจตัดสินใจ กำหนดรูปแบบและกำหนดวัตถุประสงค์ในการประมวลผลข้อมูลส่วนบุคคล ได้.....(มอบหมาย/แต่งตั้ง/จ้าง/อื่น ๆ).....ให้ {result.Contract_Organization ?? ""} ในฐานะผู้ประมวลผลข้อมูลส่วนบุคคล ดำเนินการเพื่อวัตถุประสงค์ดังต่อไปนี้
    </div>
    <div class='t-16 tab3'>วัตถุประสงค์</div>
    {(conPurpose != null && conPurpose.Count > 0
                ? $"<ul class='t-16 tab3'>{string.Join("", conPurpose.Select(p => $"<li>{p.Objective_Description}</li>"))}</ul>"
                : "<div class='t-16 tab3'>- ไม่มีข้อมูลวัตถุประสงค์ -</div>")}
    <div class='t-16 tab3'>ข้อตกลง</div>
    {(conAgreement != null && conAgreement.Count > 0
                ? $"<ul class='t-16 tab3'>{string.Join("", conAgreement.Select(a => $"<li>{a.PD_Detail}</li>"))}</ul>"
                : "<div class='t-16 tab3'>- ไม่มีข้อมูลข้อตกลง -</div>")}
    <div class='t-16 tab3'>
        ด้วยเหตุนี้ ทั้งสองฝ่ายจึงตกลงจัดทำข้อตกลงฉบับนี้ และให้ถือข้อตกลงฉบับนี้เป็นส่วนหนึ่งของ {result.Master_Contract_Number ?? ""} เพื่อเป็นหลักฐานการควบคุมดูแลการประมวลผลข้อมูลส่วนบุคคลที่ สสว. มอบหมายหรือแต่งตั้งให้ {result.Contract_Organization ?? ""} ดำเนินการ อันเนื่องมาจากการดำเนินการตามหน้าที่และความรับผิดชอบตาม {result.Master_Contract_Number ?? ""} ฉบับลงวันที่ {result.Master_Contract_Sign_Date?.ToString("dd/MM/yyyy") ?? ""} และเพื่อดำเนินการให้เป็นไปตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒ และกฎหมายอื่นๆ ที่ออกตามความในพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒ ซึ่งต่อไปในข้อตกลงฉบับนี้ รวมเรียกว่า “กฎหมายคุ้มครองข้อมูลส่วนบุคคล” ทั้งที่มีผลใช้บังคับอยู่ ณ วันทำข้อตกลงฉบับนี้และที่จะมีการเพิ่มเติมหรือแก้ไขเปลี่ยนแปลงในภายหลัง โดยมีรายละเอียดดังนี้
    </div>
    <div class='t-16 tab3'>
        1. {result.Contract_Organization ?? ""} รับทราบว่า ข้อมูลส่วนบุคคล หมายถึง ข้อมูลเกี่ยวกับบุคคลธรรมดาซึ่งทำให้สามารถระบุตัวบุคคลนั้นได้ไม่ว่าทางตรงหรือทางอ้อม โดย {result.Contract_Organization ?? ""} จะดำเนินการ ตามที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลกำหนด เพื่อคุ้มครองให้การประมวลผลข้อมูลส่วนบุคคลเป็นไปอย่างเหมาะสมและถูกต้องตามกฎหมาย
    </div>
    <div class='t-16 tab3'>
        โดยในการดำเนินการตามข้อตกลงนี้ {result.Contract_Organization ?? ""} จะประมวลผลข้อมูลส่วนบุคคลเมื่อได้รับคำสั่งที่เป็นลายลักษณ์อักษรจาก สสว. แล้วเท่านั้น ทั้งนี้ เพื่อให้ปราศจากข้อสงสัย การดำเนินการประมวลผลข้อมูลส่วนบุคคลโดย {result.Contract_Organization ?? ""} ตามหน้าที่และความรับผิดชอบตาม {result.Master_Contract_Number ?? ""} ถือเป็นการได้รับคำสั่งที่เป็นลายลักษณ์อักษรจาก สสว. แล้ว
    </div>
    <div class='t-16 tab3'>
        2. {result.Contract_Organization ?? ""} จะกำหนดให้การเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้ถูกจำกัดเฉพาะเจ้าหน้าที่ และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ได้รับมอบหมาย มีหน้าที่เกี่ยวข้องหรือมีความจำเป็นในการเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้เท่านั้น และจะดำเนินการเพื่อให้พนักงาน และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ได้รับมอบหมายจาก {result.Contract_Organization ?? ""} ทำการประมวลผลและรักษาความลับของข้อมูลส่วนบุคคลด้วยมาตรฐานเดียวกัน
    </div>
    <div class='t-16 tab3'>
        3. {result.Contract_Organization ?? ""} จะควบคุมดูแลให้เจ้าหน้าที่ และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ปฏิบัติหน้าที่ในการประมวลผลข้อมูลส่วนบุคคล ปฏิบัติตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลอย่างเคร่งครัด และดำเนินการประมวลผลข้อมูลส่วนบุคคลตามวัตถุประสงค์ของการดำเนินการตามข้อตกลงฉบับนี้เท่านั้น โดยจะไม่ทำซ้ำ คัดลอก ทำสำเนา บันทึกภาพข้อมูลส่วนบุคคลไม่ว่าทั้งหมดหรือแต่บางส่วนเป็นอันขาด เว้นแต่เป็นไปตามเงื่อนไขของบันทึกความร่วมมือหรือสัญญา หรือกฎหมายที่เกี่ยวข้องจะระบุหรือบัญญัติไว้เป็นประการอื่น
    </div>
    <!-- Continue with other contract sections as needed -->
    <div class='t-16 tab3'>
        4. {result.Contract_Organization ?? ""} จะดำเนินการเพื่อช่วยเหลือหรือสนับสนุน สสว. ในการตอบสนองต่อคำร้องที่เจ้าของข้อมูลส่วนบุคคลแจ้งต่อ สสว. อันเป็นการใช้สิทธิของเจ้าของข้อมูลส่วนบุคคลตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลในส่วนที่เกี่ยวข้องกับการประมวลผลข้อมูลส่วนบุคคลในขอบเขตของข้อตกลงฉบับนี้
    </div>
    <div class='t-16 tab3'>
        อย่างไรก็ดี ในกรณีที่เจ้าของข้อมูลส่วนบุคคลยื่นคำร้องขอใช้สิทธิดังกล่าวต่อ {result.Contract_Organization ?? ""} โดยตรง {result.Contract_Organization ?? ""} จะดำเนินการแจ้งและส่งคำร้องดังกล่าวให้แก่ สสว. ทันที โดย {result.Contract_Organization ?? ""} จะไม่เป็นผู้ตอบสนองต่อคำร้องดังกล่าว เว้นแต่ สสว. จะได้มอบหมายให้ {result.Contract_Organization ?? ""} ดำเนินการเฉพาะเรื่องที่เกี่ยวข้องกับคำร้องดังกล่าว
    </div>
    <div class='t-16 tab3'>
        5. {result.Contract_Organization ?? ""} จะจัดทำและเก็บรักษาบันทึกรายการของกิจกรรมการประมวลผลข้อมูลส่วนบุคคล (Record of Processing Activities) ทั้งหมดที่ {result.Contract_Organization ?? ""} ประมวลผลในขอบเขตของข้อตกลงฉบับนี้ และจะดำเนินการส่งมอบบันทึกรายการดังกล่าวให้แก่ สสว. ทุก {result.RetentionPeriodDays?.ToString() ?? ""} วัน และ/หรือทันทีที่ สสว. ร้องขอ
    </div>
    <div class='t-16 tab3'>
        6. {result.Contract_Organization ?? ""} จะจัดให้มีและคงไว้ซึ่งมาตรการรักษาความปลอดภัยสำหรับการประมวลผลข้อมูลที่มีความเหมาะสมทั้งในเชิงองค์กรและเชิงเทคนิคตามที่คณะกรรมการคุ้มครองข้อมูลส่วนบุคคลได้ประกาศกำหนดและ/หรือตามมาตรฐานสากล โดยคำนึงถึงลักษณะ ขอบเขต และวัตถุประสงค์ของการประมวลผลข้อมูลตามที่กำหนดในข้อตกลงฉบับนี้เป็นสำคัญ เพื่อคุ้มครองข้อมูลส่วนบุคคลจากความเสี่ยงอันเกี่ยวเนื่องกับการประมวลผลข้อมูลส่วนบุคคล เช่น ความเสียหายอันเกิดจากการละเมิด อุบัติเหตุ การลบ ทำลาย สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้ เปิดเผยหรือโอนข้อมูลส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย เป็นต้น
    </div>
    <div class='t-16 tab3'>
        7. เว้นแต่กฎหมายที่เกี่ยวข้องจะบัญญัติไว้เป็นประการอื่น {result.Contract_Organization ?? ""} จะทำการลบหรือทำลายข้อมูลส่วนบุคคลที่ทำการประมวลผลภายใต้ข้อตกลงฉบับนี้ภายใน {result.RetentionPeriodDays?.ToString() ?? ""} วัน นับแต่วันที่ดำเนินการประมวลผลเสร็จสิ้น หรือวันที่ สสว. และ {result.Contract_Organization ?? ""} ได้ตกลงเป็นลายลักษณ์อักษรให้ยกเลิก {result.Master_Contract_Number ?? ""} แล้วแต่กรณีใดจะเกิดขึ้นก่อน
    </div>
    <div class='t-16 tab3'>
        นอกจากนี้ ในกรณีปรากฏว่า {result.Contract_Organization ?? ""} หมดความจำเป็นจะต้องเก็บรักษาข้อมูลส่วนบุคคลตามข้อตกลงฉบับนี้ก่อนสิ้นระยะเวลาตามวรรคหนึ่ง {result.Contract_Organization ?? ""} จะทำการลบหรือทำลายข้อมูลส่วนบุคคลตามข้อตกลงฉบับนี้ทันที
    </div>
    <div class='t-16 tab3'>
        8. กรณีที่ {result.Contract_Organization ?? ""} พบพฤติการณ์ใด ๆ ที่มีลักษณะที่กระทบต่อการรักษาความปลอดภัยของข้อมูลบุคคลที่ {result.Contract_Organization ?? ""} ประมวลผลภายใต้ข้อตกลงฉบับนี้ ซึ่งอาจก่อให้เกิดความเสียหายจากการละเมิด อุบัติเหตุ การลบ ทำลาย สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้ เปิดเผยหรือโอนข้อมูลส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย แล้ว {result.Contract_Organization ?? ""} จะดำเนินการแจ้งให้ สสว. ทราบโดยทันทีภายในเวลาไม่เกิน ... ชั่วโมง
    </div>
    <div class='t-16 tab3'>
        9. การแจ้งถึงเหตุการละเมิดข้อมูลส่วนบุคคลที่เกิดขึ้นภายใต้ข้อตกลงนี้ {result.Contract_Organization ?? ""} จะใช้มาตรการตามที่เห็นสมควรในการระบุถึงสาเหตุของการละเมิด และป้องกันปัญหาดังกล่าวมิให้เกิดซ้ำ และจะให้ข้อมูลแก่ สสว. ภายใต้ขอบเขตที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลได้กำหนด ดังต่อไปนี้
        <ul>
            <li>รายละเอียดของลักษณะและผลกระทบที่อาจเกิดขึ้นของการละเมิด</li>
            <li>มาตรการที่ถูกใช้เพื่อลดผลกระทบของการละเมิด</li>
            <li>ข้อมูลอื่น ๆ เกี่ยวข้องกับการละเมิด</li>
        </ul>
    </div>
    <div class='t-16 tab3'>
        10. หน้าที่และความรับผิดของ {result.Contract_Organization ?? ""} ในการปฏิบัติตามข้อตกลงจะสิ้นสุดลงนับแต่วันที่ปฏิบัติงานที่ตกลงเสร็จสิ้น หรือ วันที่ {result.End_Date?.ToString("dd/MM/yyyy") ?? ""} และ สสว. ได้ตกลงเป็นลายลักษณ์อักษรให้ยกเลิก {result.Master_Contract_Number ?? ""} แล้วแต่กรณีใดจะเกิดขึ้นก่อน อย่างไรก็ดี การสิ้นผลลงของข้อตกลงนี้ ไม่กระทบต่อหน้าที่ของ {result.Contract_Organization ?? ""} ในการลบหรือทำลายข้อมูลส่วนบุคคลตามที่ได้กำหนดในข้อ 7 ของข้อตกลงฉบับนี้
    </div>
    <div class='t-16 tab3'>
        ทั้งสองฝ่ายได้อ่านและเข้าใจข้อความโดยละเอียดตลอดแล้ว เพื่อเป็นหลักฐานแห่งการนี้ ทั้งสองฝ่ายจึงได้ลงนามไว้เป็นหลักฐานต่อหน้าพยาน ณ วัน เดือน ปี ที่ระบุข้างต้น
    </div>
    <table class='signature-table'>
        <tr>
            <td>(ลงชื่อ)....................................................<br/>(....................................................)<br/>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</td>
            <td>(ลงชื่อ)....................................................<br/>(....................................................)<br/>{result.Contract_Organization ?? "ชื่อคู่สัญญา"}</td>
        </tr>
        <tr>
            <td>(ลงชื่อ)....................................................พยาน<br/>(....................................................)</td>
            <td>(ลงชื่อ)....................................................พยาน<br/>(....................................................)</td>
        </tr>
    </table>
</body>
</html>
";

            var doc = new DinkToPdf.HtmlToPdfDocument()
            {
                GlobalSettings = {
                PaperSize = DinkToPdf.PaperKind.A4,
                Orientation = DinkToPdf.Orientation.Portrait,
                Margins = new DinkToPdf.MarginSettings
                {
                    Top = 20,
                    Bottom = 20,
                    Left = 20,
                    Right = 20
                }
            },
                Objects = {
                new DinkToPdf.ObjectSettings() {
                    HtmlContent = html,
                    FooterSettings = new DinkToPdf.FooterSettings
                    {
                        FontName = "THSarabunNew",
                        FontSize = 6,
                        Line = false,
                        Center = "[page] / [toPage]"
                    }
                }
            }
            };

            var pdfBytes = _pdfConverter.Convert(doc);
            return pdfBytes;
        }
        catch
        {
            return Array.Empty<byte>();
        }
    }



    #endregion 4.1.1.2.4.บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล

    //    public async Task<byte[]> OnGetWordContact_PersernalProcessService_HtmlToPDF(string id)
    //    {
    //        var result = await _eCon.GetPDPAAsync(id);

    //        if (result == null)
    //        {
    //            throw new Exception("PDPA data not found.");
    //        }

    //        // Logo
    //        var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
    //        string logoBase64 = "";
    //        if (System.IO.File.Exists(logoPath))
    //        {
    //            var bytes = System.IO.File.ReadAllBytes(logoPath);
    //            logoBase64 = Convert.ToBase64String(bytes);
    //        }

    //        // Font
    //        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");

    //        // Objectives
    //        var objectivesList = await _eCon.GetPDPA_ObjectivesAsync(id);
    //        var agreementList = await _eCon.GetPDPA_AgreementListAsync(id);

    //        var html = $@"
    //<html>
    //<head>
    //    <meta charset='utf-8'>
    //    <style>
    //        @font-face {{
    //            font-family: 'THSarabunNew';
    //            src: url('file:///{fontPath}') format('truetype');
    //            font-weight: normal;
    //            font-style: normal;
    //        }}
    //        body {{
    //            font-family: 'THSarabunNew', 'TH SarabunPSK', 'Sarabun', sans-serif;
    //            font-size: 32pt;
    //        }}
    //        .logo {{ text-align: left; margin-top: 40px; }}
    //        .title {{ text-align: center; font-size: 44pt; font-weight: bold; margin-top: 40px; }}
    //        .subtitle {{ text-align: center; font-size: 36pt; font-weight: bold; margin-top: 20px; }}
    //        .contract {{ margin-top: 20px; font-size: 28pt; text-indent: 2em; }}
    //        .section {{ margin-top: 30px; font-size: 32pt; font-weight: bold; }}
    //        .signature-table {{ width: 100%; margin-top: 60px; font-size: 28pt; }}
    //        .signature-table td {{ text-align: center; vertical-align: top; padding: 20px; }}
    //    </style>
    //</head>
    //<body>
    //    <table style='width:100%; border-collapse:collapse; margin-top:40px;'>
    //        <tr>
    //            <td style='width:60%; text-align:left; vertical-align:top;'>
    //                <img src='data:image/jpeg;base64,{logoBase64}' width='240' height='80' />
    //            </td>
    //            <td style='width:40%; text-align:center; vertical-align:top;'>
    //                <div style='display:inline-block; border:2px solid #333; padding:20px; font-size:32pt;'>
    //                    โลโก้<br/>หน่วยงาน
    //                </div>
    //            </td>
    //        </tr>
    //    </table>
    //    <div class='title'>ข้อตกลงการประมวลผลข้อมูลส่วนบุคคล</div>
    //    <div class='subtitle'>(Data Processing Agreement)</div>
    //    <div class='title'>โครงการ {result.Objectives ?? ""}</div>
    //    <div class='subtitle'>ระหว่าง</div>
    //    <div class='subtitle'>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม กับ {result.Objectives_Other ?? ""}</div>
    //    <div class='contract'>
    //        ข้อตกลงนี้จัดทำโดย {result.CreateBy ?? ""} และปรับปรุงโดย {result.UpdateBy ?? ""} รหัสคำขอ {result.Request_ID ?? ""}
    //    </div>
    //    <div class='section'>วัตถุประสงค์</div>
    //    {(objectivesList != null && objectivesList.Count > 0 ? $"<ul>{string.Join("", objectivesList.Select((o, i) => $"<li>{o.PDPA_ID}</li>"))}</ul>" : "<div class='contract'>- ไม่มีข้อมูลวัตถุประสงค์ -</div>")}
    //    <div class='section'>ข้อตกลง</div>
    //    {(agreementList != null && agreementList.Count > 0 ? $"<ul>{string.Join("", agreementList.Select((a, i) => $"<li>{a.PDPA_ID}</li>"))}</ul>" : "<div class='contract'>- ไม่มีข้อมูลข้อตกลง -</div>")}
    //    <table class='signature-table'>
    //        <tr>
    //            <td>(ลงชื่อ)....................................................<br/>(....................................................)<br/>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</td>
    //            <td>(ลงชื่อ)....................................................<br/>(....................................................)<br/>ชื่อคู่สัญญา</td>
    //        </tr>
    //        <tr>
    //            <td>(ลงชื่อ)....................................................พยาน<br/>(....................................................)</td>
    //            <td>(ลงชื่อ)....................................................พยาน<br/>(....................................................)</td>
    //        </tr>
    //    </table>
    //</body>
    //</html>
    //";

    //        // You need to inject IConverter _pdfConverter in the constructor for PDF generation
    //        var doc = new DinkToPdf.HtmlToPdfDocument()
    //        {
    //            GlobalSettings = {
    //            PaperSize = DinkToPdf.PaperKind.A4,
    //            Orientation = DinkToPdf.Orientation.Portrait,
    //            Margins = new DinkToPdf.MarginSettings
    //            {
    //                Top = 20,
    //                Bottom = 20,
    //                Left = 20,
    //                Right = 20
    //            }
    //        },
    //            Objects = {
    //            new DinkToPdf.ObjectSettings() {
    //                HtmlContent = html,
    //                FooterSettings = new DinkToPdf.FooterSettings
    //                {
    //                    FontName = "THSarabunNew",
    //                    FontSize = 6,
    //                    Line = false,
    //                    Center = "[page] / [toPage]"
    //                }
    //            }
    //        }
    //        };

    //        var pdfBytes = _w._pdfConverter.Convert(doc);
    //        return pdfBytes;
    //    }
}
