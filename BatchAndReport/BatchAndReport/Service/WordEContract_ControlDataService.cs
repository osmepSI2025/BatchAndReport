using BatchAndReport.DAO;
using BatchAndReport.Models;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading.Tasks;

public class WordEContract_ControlDataService
{
    private readonly WordServiceSetting _w;
    private readonly E_ContractReportDAO _eContractReportDAO;
    private readonly IConverter _pdfConverter;
    public WordEContract_ControlDataService(WordServiceSetting ws
          , E_ContractReportDAO eContractReportDAO
         , IConverter pdfConverter
        )
    {
        _w = ws;
        _eContractReportDAO = eContractReportDAO;
        _pdfConverter = pdfConverter;
    }
    #region 4.1.1.2.5.บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วมตัวอย่างหน้าจอ
    public async Task<byte[]> OnGetWordContact_ControlDataService(string id)
    {
        var result = await _eContractReportDAO.GetJDCAAsync(id);
        if (result == null)
        {
            throw new Exception("ไม่พบข้อมูลสัญญา");
        }
        else {
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
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("ข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วม", "32"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("(Joint Controller Agreement)", "32"));

                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("ระหว่าง", "32"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.)", "32"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("กับ", "32"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph(result.Contract_Party_Name, "36"));


                // --- 3. Main contract body ---

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วม (“ข้อตกลง”) ฉบับนี้ ทำขึ้นเมื่อวันที่ "+result.Master_Contract_Sign_Date.ToString()+" ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("โดยที่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สสว.” ฝ่ายหนึ่ง ได้ตกลงใน "+result.MOU_Name+" ฉบับลงวันที่ "+result.Master_Contract_Sign_Date??""+" ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สัญญาหลัก” กับ  "+ result.Contract_Party_Name + "  ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “(ชื่อเรียกคู่สัญญา)” อีกฝ่ายหนึ่ง รวมทั้งสองฝ่ายว่า “คู่สัญญา”", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("เพื่อให้บรรลุตามวัตถุประสงค์ที่คู่สัญญาได้ตกลงกันภายใต้สัญญาหลัก คู่สัญญามีความจำเป็นต้องร่วมกันเก็บ รวบรวม ใช้ หรือเปิดเผย (รวมเรียกว่า “ประมวลผล”) ข้อมูลส่วนบุคคลตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. 2562 โดยที่คู่สัญญามีอำนาจตัดสินใจ กำหนดรูปแบบ รวมถึงวัตถุประสงค์ในการประมวลผลข้อมูลส่วนบุคคลนั้นร่วมกัน ในลักษณะของผู้ควบคุมข้อมูลส่วนบุคคลร่วม", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("คู่สัญญาจึงตกลงจัดทำข้อตกลงฉบับนี้ และให้ถือเป็นส่วนหนึ่งของสัญญาหลัก เพื่อกำหนดขอบเขตอำนาจหน้าที่และความรับผิดชอบของคู่สัญญาในการร่วมกันประมวลผลข้อมูลส่วนบุคคล โดยข้อตกลงนี้ใช้บังคับกับกิจกรรมการประมวลผลข้อมูลส่วนบุคคลทั้งสิ้นที่ดำเนินการโดยคู่สัญญา รวมถึงผู้ประมวลผลข้อมูลส่วนบุคคลซึ่งถูกหรืออาจถูกมอบหมายให้ประมวลผลข้อมูลส่วนบุคคลโดยคู่สัญญา ทั้งนี้ เพื่อดำเนินการให้เป็นไปตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. 2562 รวมถึงกฎหมายอื่น ๆ ที่ออกตามความของพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. 2562 ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “กฎหมายคุ้มครองข้อมูลส่วนบุคคล” ทั้งที่มีผลใช้บังคับอยู่ ณ วันที่ทำข้อตกลงฉบับนี้ และที่อาจมีเพิ่มเติมหรือแก้ไขเปลี่ยนแปลงในภายหลัง โดยมีรายละเอียดดังต่อไปนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ภายใต้ข้อตกลงของการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วมนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 1 วัตถุประสงค์และวิธีการประมวลผล", null, "32", true));
                body.AppendChild(WordServiceSetting.JustifiedParagraph("คู่สัญญาร่วมกันกำหนดวัตถุประสงค์และวิธีการในการประมวลผลข้อมูลดังรายการกิจกรรมการประมวลผลข้อมูลส่วนบุคคล (“กิจกรรมการประมวลผลข้อมูลส่วนบุคคลหลัก”) ดังต่อไปนี้ (ระบุวัตถุประสงค์ตามสัญญาหลักที่คู่สัญญาจะต้องดำเนินการร่วมกัน)", "32"));
              
                var purplist = await _eContractReportDAO.GetJDCA_JointPurpAsync(id);
                if (purplist != null && purplist.Count > 0)
                {
                    foreach (var item in purplist)
                    {
                        body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs(item.Objective_Description, null, "32"));
                    }
                }
        
                
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ซึ่งจากรายการกิจกรรมการประมวลผลหลักที่คู่สัญญาร่วมกันกำหนดวัตถุประสงค์ข้างต้น คู่สัญญาแต่ละฝ่ายมีการประมวลผลข้อมูลส่วนบุคคล (“กิจกรรมการประมวลผลข้อมูลส่วนบุคคลย่อย”) ดังรายละเอียดต่อไปนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(1) กิจกรรมการประมวลผลข้อมูลส่วนบุคคลย่อยที่ดำเนินการโดย สสว.", null, "32", true));
                // With a table that matches your screenshot:

                var dtActivitySME = await _eContractReportDAO.GetJDCA_SubProcessActivitiesAsync(id);

                if (dtActivitySME != null && dtActivitySME.Count != 0)
                {
                   var activityListOSMEP = dtActivitySME.Where(x => x.Owner == "OSMEP").ToList();
                    var activityListCP = dtActivitySME.Where(x => x.Owner == "CP").ToList();
                    foreach(var item in activityListOSMEP)
                    {
                        var tableActivitySME = new Table(
         new TableProperties(
           new TableWidth { Width = "9000", Type = TableWidthUnitValues.Dxa }, // 9000 twips fits A4 with margins
           new TableBorders(
             new TopBorder { Val = BorderValues.Single },
             new BottomBorder { Val = BorderValues.Single },
             new LeftBorder { Val = BorderValues.Single },
             new RightBorder { Val = BorderValues.Single },
             new InsideHorizontalBorder { Val = BorderValues.Single },
             new InsideVerticalBorder { Val = BorderValues.Single }
           )
         ),
         new TableRow(
           new TableCell(
             new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }),
             WordServiceSetting.CenteredBoldParagraph("รายการกิจกรรมการประมวลผล")
           ),
           new TableCell(
             new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }),
             WordServiceSetting.CenteredBoldParagraph("ฐานกฎหมายที่ใช้ในการประมวลผล")
           ),
           new TableCell(
             new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }),
             WordServiceSetting.CenteredBoldParagraph("รายการข้อมูลส่วนบุคคลที่ใช้ประมวลผล")
           )
         ),
         new TableRow(
           new TableCell(
             new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }),
             WordServiceSetting.JustifiedParagraph(
              item.Activity ?? "", "32")
           ),
           new TableCell(
             new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }),
             WordServiceSetting.JustifiedParagraph(
             item.LegalBasis ?? "", "32")
           ),
           new TableCell(
             new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }),
             WordServiceSetting.JustifiedParagraph(
               item.PersonalData ?? "", "32")
           )
         )
       );
                        body.AppendChild(tableActivitySME);

                    }




                    body.AppendChild(WordServiceSetting.EmptyParagraph());
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(2) กิจกรรมการประมวลผลข้อมูลส่วนบุคคลย่อยซึ่งดำเนินการโดย (" + result.Contract_Party_Name + ")", null, "32", true));

                    foreach (var item in activityListOSMEP)
                    {
                        var tableActivityCustomer = new Table(
                 new TableProperties(
                   new TableWidth { Width = "9000", Type = TableWidthUnitValues.Dxa }, // 9000 twips fits A4 with margins
                   new TableBorders(
                     new TopBorder { Val = BorderValues.Single },
                     new BottomBorder { Val = BorderValues.Single },
                     new LeftBorder { Val = BorderValues.Single },
                     new RightBorder { Val = BorderValues.Single },
                     new InsideHorizontalBorder { Val = BorderValues.Single },
                     new InsideVerticalBorder { Val = BorderValues.Single }
                   )
                 ),
                 new TableRow(
                   new TableCell(
                     new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }),
                     WordServiceSetting.CenteredBoldParagraph("รายการกิจกรรมการประมวลผล")
                   ),
                   new TableCell(
                     new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }),
                     WordServiceSetting.CenteredBoldParagraph("ฐานกฎหมายที่ใช้ในการประมวลผล")
                   ),
                   new TableCell(
                     new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }),
                     WordServiceSetting.CenteredBoldParagraph("รายการข้อมูลส่วนบุคคลที่ใช้ประมวลผล")
                   )
                 ),

                 new TableRow(
                   new TableCell(
                     new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }),
                     WordServiceSetting.JustifiedParagraph(
                             item.Activity ?? "", "32")
                   ),
                   new TableCell(
                     new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }),
                     WordServiceSetting.JustifiedParagraph(
                             item.LegalBasis ?? "", "32")
                   ),
                   new TableCell(
                     new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }),
                     WordServiceSetting.JustifiedParagraph(
                             item.PersonalData ?? "", "32")
                   )
                 )
               );
                        body.AppendChild(tableActivityCustomer);
                    }
                   
                }
           
          
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ทั้งนี้ คู่สัญญาแต่ละฝ่ายรับรองว่าจะดำเนินการประมวลผลข้อมูลส่วนบุคคลดังรายละเอียดข้างต้นให้เป็นไปตามที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลกำหนด" +
                    " โดยเฉพาะอย่างยิ่งในเรื่องความชอบด้วยกฎหมายของการประมวลผลข้อมูลภายใต้ความเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วม" +
                    " โดยคู่สัญญาแต่ละฝ่ายจะจัดให้มีและคงไว้ซึ่งมาตรการรักษาความปลอดภัยสำหรับการประมวลผลข้อมูลที่มีความเหมาะสมทั้งในมาตรการเชิงองค์กร มาตรการเชิงเทคนิค" +
                    " และมาตรการเชิงกายภาพ ตามที่คณะกรรมการคุ้มครองข้อมูลส่วนบุคคลได้ประกาศกำหนดและ/หรือตามมาตรฐานสากล" +
                    " โดยคำนึงถึงลักษณะ ขอบเขต และวัตถุประสงค์ของการประมวลผลข้อมูล เพื่อคุ้มครองข้อมูลส่วนบุคคลจากความเสี่ยงอันเกี่ยวเนื่องกับการประมวลผลข้อมูลส่วนบุคคล" +
                    " เช่น ความเสียหายอันเกิดจากการละเมิด อุบัติเหตุ การลบ ทำลาย สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้ เปิดเผยหรือโอนข้อมูลส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย เป็นต้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 2 หน้าที่และความรับผิดชอบของคู่สัญญา", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("2.1 คู่สัญญารับรองว่าจะควบคุมดูแลให้เจ้าหน้าที่ พนักงาน และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ปฏิบัติหน้าที่ในการประมวลผล" +
                    "ข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้รักษาความลับและปฏิบัติตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลอย่างเคร่งครัด และดำเนินการประมวลผลข้อมูลส่วนบุคคลเพื่อวัตถุประสงค์ตามข้อตกลงฉบับนี้" +
                    "เท่านั้น โดยจะไม่ทำซ้ำ คัดลอก ทำสำเนา บันทึกภาพข้อมูลส่วนบุคคลไม่ว่าทั้งหมดหรือแต่บางส่วนเป็นอันขาด เว้นแต่ เป็นไปตามเงื่อนไขของสัญญาหลัก หรือกฎหมายที่เกี่ยวข้องจะระบุหรือบัญญัติไว้เป็นประการอื่น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("2.2 คู่สัญญารับรองว่าจะกำหนดให้การเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้ถูกจำกัดเฉพาะเจ้าหน้าที่ พนักงาน และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ได้รับมอบหมาย มีหน้าที่เกี่ยวข้องหรือมีความจำเป็นในการเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้เท่านั้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("2.3 คู่สัญญาจะไม่เปิดเผยข้อมูลส่วนบุคคลภายใต้ข้อตกลงนี้แก่บุคคลที่ไม่มีอำนาจหน้าที่เกี่ยวข้องในการประมวลผล หรือบุคคลภายนอก เว้นแต่ กรณีที่มีความจำเป็นต้องกระทำนตามหน้าที่ในสัญญาหลัก ของข้อตกลงฉบับนี้ หรือเพื่อปฏิบัติตามกฎหมายที่ใช้บังคับ หรือที่ได้รับความยินยอมจากคู่สัญญาอีกฝ่ายก่อน", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("2.4 คู่สัญญาแต่ละฝ่ายมีหน้าที่ต้องแจ้งรายละเอียดของการประมวลผลข้อมูลส่วนบุคคลแก่เจ้าของข้อมูลส่วนบุคคลซึ่งถูกประมวลผลข้อมูลก่อนหรือขณะเก็บรวบรวมข้อมูล" +
                    "ส่วนบุคคล ทั้งนี้รายการรายละเอียดที่ต้องแจ้งให้เป็นไปตามที่กำหนดในมาตรา 23 แห่งพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. 2562", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("2.5 กรณีที่คู่สัญญาฝ่ายหนึ่งฝ่ายใด พบพฤติการณ์ที่มีลักษณะที่กระทบต่อการรักษาความปลอดภัยของข้อมูลส่วนบุคคลที่ประมวลผลภายใต้ข้อตกลงฉบับนี้ ซึ่งอาจก่อให้เกิดความเสียหายจากการละเมิด อุบัติเหตุ การลบ ทำลาย สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้ เปิดเผยหรือโอนข้อมูลส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย คู่สัญญาฝ่ายที่พบเหตุดังกล่าวจะดำเนินการแจ้งให้คู่สัญญาอีกฝ่ายทราบพร้อมรายละเอียดของเหตุการณ์โดยไม่ชักช้าภายใน 72 ชั่วโมง นับแต่ผู้ประมวลข้อมูลทราบเหตุเท่าที่จะสามารถกระทำได้ ทั้งนี้ คู่สัญญาแต่ละฝ่ายต่างมีหน้าที่ต้องแจ้งเหตุดังกล่าวแก่สำนักงานคณะกรรมการคุ้มครองข้อมูลส่วนบุคคล หรือเจ้าของข้อมูลส่วนบุคคล ตามแต่กรณีที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลกำหนดไว้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("2.6 คู่สัญญาตกลงจะให้ความช่วยเหลืออย่างสมเหตุสมผลแก่อีกฝ่ายในการตอบสนองต่อข้อเรียกร้องใด ๆ ที่สมเหตุสมผลจากการใช้สิทธิต่าง ๆ ภายใต้กฎหมายคุ้มครองข้อมูลส่วนบุคคลโดยเจ้าของข้อมูลส่วนบุคคล โดยพิจารณาถึงลักษณะการประมวลผล ภาระหน้าที่ภายใต้กฎหมายคุ้มครองข้อมูลที่ใช้บังคับ และข้อมูลส่วนบุคคลที่ประมวลผล ทั้งนี้ คู่สัญญาทราบว่าเจ้าของข้อมูลส่วนบุคคลอาจยื่นคำร้องขอใช้สิทธิดังกล่าวต่อคู่สัญญาฝ่ายหนึ่งฝ่ายใดก็ได้ ซึ่งคู่สัญญาฝ่ายที่ได้รับคำร้องจะต้องดำเนินการแจ้งถึงคำร้องดังกล่าวแก่คู่สัญญาอีกฝ่ายโดยทันที โดยคู่สัญญาฝ่ายที่รับคำร้องนั้นจะต้องแจ้งให้เจ้าของข้อมูลทราบถึงการจัดการตามคำขอหรือข้อร้องเรียนของเจ้าของข้อมูลนั้นด้วย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("2.7 ในกรณีที่มีการใช้ผู้ประมวลผลข้อมูลส่วนบุคคลเพื่อทำการประมวลผลข้อมูลส่วนบุคคลภายใต้ข้อตกลงนี้ ให้ดำเนินการแจ้งต่อคู่สัญญาอีกฝ่ายก่อน ทั้งนี้ คู่สัญญาฝ่ายที่ใช้ผู้ประมวลผลข้อมูลส่วนบุคคลจะต้องทำสัญญากับผู้ประมวลผลข้อมูลเป็นลายลักษณ์อักษรตามเงื่อนไขที่กฎหมายคุ้มครองข้อมูลกำหนด เพื่อหลีกเลี่ยงข้อสงสัย หากคู่สัญญาฝ่ายหนึ่งฝ่ายใดได้ว่าจ้างหรือมอบหมายผู้ประมวลผลข้อมูลส่วนบุคคล คู่สัญญาฝ่ายนั้นยังคงต้องมีความรับผิดต่ออีกฝ่ายสำหรับการกระทำการหรือละเว้นกระทำการใด ๆ ของผู้ประมวลผลข้อมูลส่วนบุคคลนั้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3 การชดใช้ค่าเสียหาย", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.1 คู่สัญญาแต่ละฝ่ายจะต้องชดใช้ความเสียหายให้แก่อีกฝ่ายในค่าปรับ ความสูญหายหรือเสียหายใด ๆ ที่เกิดขึ้นกับฝ่ายที่ไม่ได้ผิดเงื่อนไข อันเนื่องมาจากการฝ่าฝืนข้อตกลงฉบับนี้ แม้ว่าจะมีข้อจำกัดความรับผิดภายใต้สัญญาหลักก็ตาม", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(1) สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.) ร้อยละ 50", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ทั้งนี้การตกลงกันของคู่สัญญานี้ ไม่มีอำนาจเหนือไปกว่าคำพิพากษาหรือคำสั่งถึงที่สุดของศาลหรือหน่วยงานผู้มีอำนาจที่กำหนดให้คู่สัญญาหรือคู่สัญญาฝ่ายหนึ่งฝ่ายใดต้องถูกปรับหรือชดใช้ค่าเสียหาย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("4 ระยะเวลาตามข้อตกลง", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("หน้าที่และความรับผิดของคู่สัญญาในการปฏิบัติตามข้อตกลงฉบับนี้จะสิ้นสุดลงนับแต่วันที่การดำเนินการตามสัญญาหลักเสร็จสิ้นลง หรือ วันที่คู่สัญญาได้ตกลงเป็นลายลักษณ์อักษรให้ยกเลิกสัญญาหลัก แล้วแต่กรณีใดจะเกิดขึ้นก่อน", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("5 ผู้แทนของคู่สัญญาแต่ละฝ่าย", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("คู่สัญญาตกลงแต่งตั้งผู้แทนของแต่ละฝ่าย ดังรายการต่อไปนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(1) สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.)", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้แทน :"+result.OSMEP_ContRep+"", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ติดต่อได้ที่ :" + result.OSMEP_ContRep_Contact + "", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เจ้าหน้าที่คุ้มครองข้อมูลส่วนบุคคล : "+result.OSMEP_DPO+"", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ติดต่อได้ที่ :" + result.OSMEP_DPO_Contact + "", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(2)"+result.Contract_Party_Name+"", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้แทน :" + result.CP_ContRep + "", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ติดต่อได้ที่ :" + result.CP_ContRep_Contact + "", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เจ้าหน้าที่คุ้มครองข้อมูลส่วนบุคคล (ถ้ามี) :" + result.CP_DPO + "", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ติดต่อได้ที่ :" + result.CP_DPO_Contact + "", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("6.การบังคับใช้", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ในกรณีที่ข้อตกลง คำรับรอง การเจรจาหรือข้อผูกพันใดที่คู่สัญญามีต่อกันไม่ว่าด้วยวาจาหรือเป็นลายลักษณ์อักษรก็ดี ขัดหรือแย้งกับข้อความที่ระบุในข้อตกลงฉบับนี้ ให้ใช้ข้อความตามข้อตกลงฉบับนี้บังคับ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อตกลงฉบับนี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน คู่สัญญาทั้งสองฝ่ายได้อ่าน และเข้าใจข้อความในข้อตกลงโดยละเอียดตลอดแล้ว เห็นว่าตรงตามเจตนารมณ์ทุกประการ เพื่อเป็นหลักฐานแห่งการนี้ ทั้งสองฝ่ายจึงได้ลงลายมือชื่อพร้อมทั้งประทับตราสำคัญผูกพันนิติบุคคล (ถ้ามี) ไว้เป็นหลักฐานณ วัน เดือน ปี ที่ระบุข้างต้น และคู่สัญญาต่างยึดถือไว้ฝ่ายละหนึ่งฉบับ", null, "32"));

                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                // --- 6. Signature lines ---
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // Main signature table: ผู้ให้ข้อมูล (left) | ผู้รับข้อมูล (right)
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
                            WordServiceSetting.RightParagraph("ลงชื่อ.................................................................ผู้ให้ข้อมูล")
                        ),
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.RightParagraph("ลงชื่อ.................................................................ผู้รับข้อมูล")
                        )
                    ),
                    // Second row: organization names
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.CenteredParagraph("(สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม)")
                        ),
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.CenteredBoldColoredParagraph("("+ result.Contract_Party_Name + ")", "#00000")
                        )
                    )
                );
                body.AppendChild(signatureTable);
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // Witness table: พยาน (left/right)
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
                    // Second row: organization names
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.CenteredParagraph("(สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม)")
                        ),
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.CenteredBoldColoredParagraph("(" + result.Contract_Party_Name + ")", "#00000")
                        )
                    )
                );
                body.AppendChild(witnessTable);
                body.AppendChild(WordServiceSetting.EmptyParagraph());


                // --- 7. Add header/footer if needed ---
                WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);
            }
            stream.Position = 0;
            return stream.ToArray();
        }
       
    }
    public async Task<byte[]> OnGetWordContact_ControlDataServiceHtmlToPdf(string id)
    {
        var result = await _eContractReportDAO.GetJDCAAsync(id);
        if (result == null)
        {
            throw new Exception("ไม่พบข้อมูลสัญญา");
        }

        var purplist = await _eContractReportDAO.GetJDCA_JointPurpAsync(id);
        var dtActivitySME = await _eContractReportDAO.GetJDCA_SubProcessActivitiesAsync(id);

        var activityListOSMEP = dtActivitySME?.Where(x => x.Owner == "OSMEP").ToList() ?? new List<E_ConReport_JDCA_SubProcessActivitiesModels>();
        var activityListCP = dtActivitySME?.Where(x => x.Owner == "CP").ToList() ?? new List<E_ConReport_JDCA_SubProcessActivitiesModels>();

        var html = $@"
<html>
<head>
    <meta charset='utf-8'>
    <style>
        body {{ font-family: 'THSarabunNew', 'Sarabun', sans-serif; font-size: 32pt; }}
        .title {{ text-align: center; font-size: 44pt; font-weight: bold; margin-top: 40px; }}
        .subtitle {{ text-align: center; font-size: 36pt; font-weight: bold; margin-top: 20px; }}
        .section {{ margin-top: 30px; font-size: 32pt; font-weight: bold; }}
        .contract {{ margin-top: 20px; font-size: 28pt; text-indent: 2em; }}
        .table {{ width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 28pt; }}
        .table th, .table td {{ border: 1px solid #000; padding: 8px; }}
        .signature-table {{ width: 100%; margin-top: 60px; font-size: 28pt; }}
        .signature-table td {{ text-align: center; vertical-align: top; padding: 20px; }}
    </style>
</head>
<body>
    <div class='title'>ข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วม</div>
    <div class='subtitle'>(Joint Controller Agreement)</div>
    <div class='contract'>ระหว่าง</div>
    <div class='contract'>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.)</div>
    <div class='contract'>กับ {result.Contract_Party_Name ?? ""}</div>
    <div class='contract'>ข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วม (“ข้อตกลง”) ฉบับนี้ ทำขึ้นเมื่อวันที่ {result.Master_Contract_Sign_Date?.ToString("dd/MM/yyyy") ?? ""} ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</div>
    <div class='contract'>โดยที่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สสว.” ฝ่ายหนึ่ง ได้ตกลงใน {result.MOU_Name ?? ""} ฉบับลงวันที่ {result.Master_Contract_Sign_Date?.ToString("dd/MM/yyyy") ?? ""} ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สัญญาหลัก” กับ  {result.Contract_Party_Name ?? ""}  ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “(ชื่อเรียกคู่สัญญา)” อีกฝ่ายหนึ่ง รวมทั้งสองฝ่ายว่า “คู่สัญญา”</div>
    <div class='contract'>เพื่อให้บรรลุตามวัตถุประสงค์ที่คู่สัญญาได้ตกลงกันภายใต้สัญญาหลัก คู่สัญญามีความจำเป็นต้องร่วมกันเก็บ รวบรวม ใช้ หรือเปิดเผย (รวมเรียกว่า “ประมวลผล”) ข้อมูลส่วนบุคคลตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. 2562 โดยที่คู่สัญญามีอำนาจตัดสินใจ กำหนดรูปแบบ รวมถึงวัตถุประสงค์ในการประมวลผลข้อมูลส่วนบุคคลนั้นร่วมกัน ในลักษณะของผู้ควบคุมข้อมูลส่วนบุคคลร่วม</div>
    <div class='contract'>คู่สัญญาจึงตกลงจัดทำข้อตกลงฉบับนี้ และให้ถือเป็นส่วนหนึ่งของสัญญาหลัก เพื่อกำหนดขอบเขตอำนาจหน้าที่และความรับผิดชอบของคู่สัญญาในการร่วมกันประมวลผลข้อมูลส่วนบุคคล โดยข้อตกลงนี้ใช้บังคับกับกิจกรรมการประมวลผลข้อมูลส่วนบุคคลทั้งสิ้นที่ดำเนินการโดยคู่สัญญา รวมถึงผู้ประมวลผลข้อมูลส่วนบุคคลซึ่งถูกหรืออาจถูกมอบหมายให้ประมวลผลข้อมูลส่วนบุคคลโดยคู่สัญญา ทั้งนี้ เพื่อดำเนินการให้เป็นไปตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. 2562 รวมถึงกฎหมายอื่น ๆ ที่ออกตามความของพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. 2562 ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “กฎหมายคุ้มครองข้อมูลส่วนบุคคล” ทั้งที่มีผลใช้บังคับอยู่ ณ วันที่ทำข้อตกลงฉบับนี้ และที่อาจมีเพิ่มเติมหรือแก้ไขเปลี่ยนแปลงในภายหลัง โดยมีรายละเอียดดังต่อไปนี้</div>
    <div class='section'>ข้อ 1 วัตถุประสงค์และวิธีการประมวลผล</div>
    <div class='contract'>คู่สัญญาร่วมกันกำหนดวัตถุประสงค์และวิธีการในการประมวลผลข้อมูลดังรายการกิจกรรมการประมวลผลข้อมูลส่วนบุคคล (“กิจกรรมการประมวลผลข้อมูลส่วนบุคคลหลัก”) ดังต่อไปนี้ (ระบุวัตถุประสงค์ตามสัญญาหลักที่คู่สัญญาจะต้องดำเนินการร่วมกัน)</div>
    <ul>
        {string.Join("", purplist.Select(x => $"<li>{x.Objective_Description}</li>"))}
    </ul>
    <div class='contract'>ซึ่งจากรายการกิจกรรมการประมวลผลหลักที่คู่สัญญาร่วมกันกำหนดวัตถุประสงค์ข้างต้น คู่สัญญาแต่ละฝ่ายมีการประมวลผลข้อมูลส่วนบุคคล (“กิจกรรมการประมวลผลข้อมูลส่วนบุคคลย่อย”) ดังรายละเอียดต่อไปนี้</div>
    <div class='section'>(1) กิจกรรมการประมวลผลข้อมูลส่วนบุคคลย่อยที่ดำเนินการโดย สสว.</div>
    <table class='table'>
        <tr>
            <th>รายการกิจกรรมการประมวลผล</th>
            <th>ฐานกฎหมายที่ใช้ในการประมวลผล</th>
            <th>รายการข้อมูลส่วนบุคคลที่ใช้ประมวลผล</th>
        </tr>
        {string.Join("", activityListOSMEP.Select(x => $@"
        <tr>
            <td>{x.Activity}</td>
            <td>{x.LegalBasis}</td>
            <td>{x.PersonalData}</td>
        </tr>"))}
    </table>
    <div class='section'>(2) กิจกรรมการประมวลผลข้อมูลส่วนบุคคลย่อยซึ่งดำเนินการโดย ({result.Contract_Party_Name ?? ""})</div>
    <table class='table'>
        <tr>
            <th>รายการกิจกรรมการประมวลผล</th>
            <th>ฐานกฎหมายที่ใช้ในการประมวลผล</th>
            <th>รายการข้อมูลส่วนบุคคลที่ใช้ประมวลผล</th>
        </tr>
        {string.Join("", activityListCP.Select(x => $@"
        <tr>
            <td>{x.Activity}</td>
            <td>{x.LegalBasis}</td>
            <td>{x.PersonalData}</td>
        </tr>"))}
    </table>
    <!-- Add more sections as needed, following your Word structure -->
    <div class='section'>ข้อ 2 หน้าที่และความรับผิดชอบของคู่สัญญา</div>
    <!-- ... (add all paragraphs as in your Word logic) ... -->
    <div class='section'>5 ผู้แทนของคู่สัญญาแต่ละฝ่าย</div>
    <div class='contract'>(1) สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.)</div>
    <div class='contract'>ผู้แทน : {result.OSMEP_ContRep ?? ""}</div>
    <div class='contract'>ติดต่อได้ที่ : {result.OSMEP_ContRep_Contact ?? ""}</div>
    <div class='contract'>เจ้าหน้าที่คุ้มครองข้อมูลส่วนบุคคล : {result.OSMEP_DPO ?? ""}</div>
    <div class='contract'>ติดต่อได้ที่ : {result.OSMEP_DPO_Contact ?? ""}</div>
    <div class='contract'>(2) {result.Contract_Party_Name ?? ""}</div>
    <div class='contract'>ผู้แทน : {result.CP_ContRep ?? ""}</div>
    <div class='contract'>ติดต่อได้ที่ : {result.CP_ContRep_Contact ?? ""}</div>
    <div class='contract'>เจ้าหน้าที่คุ้มครองข้อมูลส่วนบุคคล (ถ้ามี) : {result.CP_DPO ?? ""}</div>
    <div class='contract'>ติดต่อได้ที่ : {result.CP_DPO_Contact ?? ""}</div>
    <table class='signature-table'>
        <tr>
            <td>ลงชื่อ.................................................................ผู้ให้ข้อมูล<br/>(สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม)</td>
            <td>ลงชื่อ.................................................................ผู้รับข้อมูล<br/>({result.Contract_Party_Name ?? ""})</td>
        </tr>
        <tr>
            <td>ลงชื่อ.................................................................พยาน<br/>(สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม)</td>
            <td>ลงชื่อ.................................................................พยาน<br/>({result.Contract_Party_Name ?? ""})</td>
        </tr>
    </table>
</body>
</html>
";

        // You must inject IConverter _pdfConverter in the constructor
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

    #endregion 4.1.1.2.5.บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วมตัวอย่างหน้าจอ
}
