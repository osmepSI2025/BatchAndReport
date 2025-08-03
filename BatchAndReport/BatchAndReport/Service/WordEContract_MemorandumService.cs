using BatchAndReport.DAO;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading.Tasks;

public class WordEContract_MemorandumService
{
    private readonly WordServiceSetting _w;
    private readonly E_ContractReportDAO _eContractReportDAO;
    private readonly IConverter _pdfConverter; 
    public WordEContract_MemorandumService(WordServiceSetting ws
            , E_ContractReportDAO eContractReportDAO
        , IConverter pdfConverter
        )
    {
        _w = ws;
        _eContractReportDAO = eContractReportDAO;
        _pdfConverter = pdfConverter;
    }
    #region 4.1.1.2.3.บันทึกข้อตกลงความร่วมมือ
    public async Task<byte[]> OnGetWordContact_MemorandumService(string id)
    {
        var result = await _eContractReportDAO.GetMOUAsync(id);

        if (result == null)
        {
            throw new Exception("ไม่พบข้อมูลบันทึกข้อตกลงความร่วมมือ");
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
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("บันทึกข้อตกลงความร่วมมือ", "44"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("โครงการ"+ result.ProjectTitle , "44"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("ระหว่าง", "32"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม", "32"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("กับ", "36"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph(result.OrgCommonName ?? "", "36"));


                // --- 3. Main contract body ---
                var strDateTH = CommonDAO.ToThaiDateStringCovert(result.Sign_Date ?? DateTime.Now);
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(บันทึกข้อตกลงความร่วมมือฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม" +
                        "เมื่อ"+ strDateTH + " ระหว่าง", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม โดย "+result.OrgCommonName+" สำนักงานตั้งอยู่เลขที่ 21 อาคารทีเอสที ทาวเวอร์ ชั้น G,17-18,23 ถนนวิภาวดีรังสิต แขวงจอมพล เขตจตุจักร กรุงเทพมหานคร 10900 ซึ่งต่อไป ในสัญญาฉบับนี้จะเรียกว่า“สสว.”ฝ่ายหนึ่ง กับ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("“ชื่อเต็มของหน่วยงาน” โดย "+result.Requestor+" ตำแหน่ง."+result.RequestorPosition+ ".ผู้มีอำนาจกระทำการแทนปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ ฉบับลง"+ strDateTH + "สำนักงานตั้งอยู่เลขที่ ซึ่งต่อไปในสัญญาฉบับนี้จะเรียกว่า “  ” อีกฝ่ายหนึ่ง", null, "32"));
                body.AppendChild(WordServiceSetting.JustifiedParagraph("วัตถุประสงค์ของความร่วมมือ", "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ทั้งสองฝ่ายมีความประสงค์ที่จะร่วมมือกันเพื่อดำเนินการภายใต้โครงการ (ชื่อโครงการที่ระบุไว้ข้างต้น) ซึ่งในบันทึกข้อตกลงฉบับนี้ต่อไปจะเรียกว่า “โครงการ” โดยมีรายละเอียดโครงการแผนการดำเนินงาน แผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) และบรรดาเอกสารแนบท้ายบันทึกข้อตกลงฉบับนี้ ซึ่งให้ถือเป็นส่วนหนึ่งของบันทึกข้อตกลงฉบับนี้ มีระยะเวลา" +
                         "ตั้งแต่วันที่ " + CommonDAO.ToThaiDateStringCovert(result.Start_Date ?? DateTime.Now) +
                    " จนถึงวันที่" +
                    CommonDAO.ToThaiDateStringCovert(result.End_Date ?? DateTime.Now) +
                    "โดยมีวัตถุประสงค์ในการดำเนินโครงการ ดังนี้",
                    null, "32"));

                var purposeList = await _eContractReportDAO.GetMOUPoposeAsync(id);
                if (purposeList.Count != 0 && purposeList != null)
                {
                    int row = 1;
                    foreach (var purpose in purposeList)
                    {
                        body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs(row.ToString() + "• " + purpose.Detail, null, "32"));
                        row++;
                    }
                }



                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 1 ขอบเขตความร่วมมือของ “สสว.”", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("1.1 ตกลงร่วมดำเนินการโครงการโดยสนับสนุนงบประมาณ " +
                                  "จำนวน " + result.Contract_Value + " บาท ( " + CommonDAO.NumberToThaiText(result.Contract_Value ?? 0) + " ) " +
                    " ซึ่งได้รวมภาษีมูลค่าเพิ่ม ตลอดจนค่าภาษีอากรอื่น ๆ แล้วให้กับ “ชื่อหน่วยร่วม”และการใช้จ่ายเงินให้เป็นไปตามแผนการจ่ายเงินตามเอกสารแนบท้ายบันทึกข้อตกลงฉบับนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("1.2 ประสานการดำเนินโครงการ เพื่อให้บรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("1.3 กำกับ ติดตามและประเมินผลการดำเนินงานของโครงการ", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 2 ขอบเขตความร่วมมือของ “ชื่อหน่วยร่วม”", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("2.1 ตกลงที่จะร่วมดำเนินการโครงการตามวัตถุประสงค์ของการโครงการและขอบเขตการดำเนินการตามรายละเอียดโครงการ แผนการดำเนินการ และแผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) ที่แนบท้ายบันทึกข้อตกลงฉบับนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("2.2 ต้องดำเนินโครงการ ปฏิบัติตามแผนการดำเนินงาน แผนการใช้จ่ายเงิน (หรืออาจมีคู่มือการดำเนินโครงการก็ได้) อย่างเคร่งครัดและให้แล้วเสร็จภายในระยะเวลาโครงการ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("2.3 ต้องประสานการดำเนินโครงการ เพื่อให้โครงการบรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("2.4 ต้องให้ความร่วมมือกับ สสว. ในการกำกับ ติดตามและประเมินผลการดำเนินงานของโครงการ", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 3 อื่น ๆ", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("3.1 หากฝ่ายใดฝ่ายหนึ่งประสงค์จะขอแก้ไข เปลี่ยนแปลง ขยายระยะเวลาของโครงการ จะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษร และต้องได้รับความยินยอมเป็น ลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และต้องทำบันทึกข้อตกลงแก้ไข เปลี่ยนแปลง ขยายระยะเวลา เพื่อลงนามยินยอมทั้งสองฝ่าย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("3.2 หากฝ่ายใดฝ่ายหนึ่งประสงค์จะขอบอกเลิกบันทึกข้อตกลงความร่วมมือก่อนครบกำหนดระยะเวลาดำเนินโครงการจะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษรไม่น้อยกว่า 30 วัน และต้องได้รับความยินยอมเป็นลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และ “ชื่อหน่วยร่วม” จะต้องคืนเงินในส่วนที่ยังไม่ได้ใช้จ่ายหรือส่วนที่เหลือทั้งหมดพร้อมดอกผล (ถ้ามี) ให้แก่ สสว. ภายใน 15 วัน นับจากวันที่ได้รับหนังสือของฝ่ายที่ยินยอมให้บอกเลิก", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("3.3 สสว. อาจบอกเลิกบันทึกข้อตกลงความร่วมมือได้ทันที หากตรวจสอบ หรือปรากฏข้อเท็จจริงว่า การใช้จ่ายเงินของ “ชื่อหน่วยร่วม” ไม่เป็นไปตามวัตถุประสงค์ของโครงการ แผนการดำเนินงาน และแผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) ทั้งมีสิทธิเรียกเงินคงเหลือพร้อมดอกผล (ถ้ามี) คืนทั้งหมดได้ทันที", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("3.4 ทรัพย์สินใด ๆ และ/หรือ สิทธิใด ๆ ที่ได้มาจากเงินสนับสนุนตามบันทึกข้อตกลงฉบับนี้ เมื่อสิ้นสุดโครงการให้ตกได้แก่ สสว. ทั้งสิ้น เว้นแต่ สสว. จะกำหนดให้เป็นอย่างอื่น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("3.5 “ชื่อหน่วยร่วม” ต้องไม่ดำเนินการในลักษณะการจ้างเหมา กับหน่วยงาน องค์กร หรือบุคคลอื่น ๆ ยกเว้นกรณีการจัดหา จัดจ้าง เป็นกิจกรรมหรือเป็นเรื่อง ๆ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("3.6 ในกรณีที่การดำเนินการตามบันทึกข้อตกลงฉบับนี้ เกี่ยวข้องกับข้อมูลส่วนบุคคลและการคุ้มครองทรัพย์สินทางปัญญา “ชื่อหน่วยร่วม” จะต้องปฏิบัติตามกฎหมายว่าด้วยการคุ้มครองข้อมูลส่วนบุคคลและการคุ้มครองทรัพย์สินทางปัญญาอย่างเคร่งครัด และหากเกิดความเสียหายหรือมีการฟ้องร้องใด ๆ “ชื่อหน่วยร่วม” จะต้องเป็นผู้รับผิดชอบต่อการละเมิดบทบัญญัติแห่งกฎหมายดังกล่าวแต่เพียงฝ่ายเดียวโดยสิ้นเชิง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("บันทึกข้อตกลงความร่วมมือฉบับนี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน ทั้งสองฝ่ายได้อ่านและเข้าใจข้อความโดยละเอียดแล้ว จึงได้ลงลายมือชื่อพร้อมประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยานและยึดถือไว้ฝ่ายละฉบับ", null, "32"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());


                // --- 6. Signature lines ---
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม
                body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ)...................................................."));
                body.AppendChild(WordServiceSetting.CenteredParagraph("("+result.OSMEP_Signer??""+")"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // หน่วยงานร่วม
                body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ)...................................................."));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(" + result.OSMEP_Witness??"" + ")"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("ชื่อเต็มหน่วยงาน"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // พยาน 1
                body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ)....................................................พยาน"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(" + result.Contract_Signer??"" + ")"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // พยาน 2
                body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ)....................................................พยาน"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(" + result.Contract_Witness??"" + ")"));
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                // --- 7. Add header/footer if needed ---
                WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);
            }
            stream.Position = 0;
            return stream.ToArray();
        }

    }

    public async Task<byte[]> OnGetWordContact_MemorandumService_HtmlToPDF(string id)
    {
        var result = await _eContractReportDAO.GetMOUAsync(id);

        if (result == null)
        {
            throw new Exception("ไม่พบข้อมูลบันทึกข้อตกลงความร่วมมือ");
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

        // Purpose list
        var purposeList = await _eContractReportDAO.GetMOUPoposeAsync(id);

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
        .tab1 {{ text-indent: 48px;  word-break: break-all;  }}
        .tab2 {{ text-indent: 96px;  word-break: break-all; }}
        .tab3 {{ text-indent: 144px;  word-break: break-all; }}
        .tab4 {{ text-indent: 192px;  word-break: break-all;}}
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
    <table class='contract-table'>
        <tr>
            <td class='logo-cell'>
                <img src='data:image/jpeg;base64,{logoBase64}' alt='Logo' style='height:80px;' />
            </td>
            <td class='code-cell'>
                <!-- Contract code box, add code if needed -->
            </td>
        </tr>
    </table>
    <div class='t-22 text-center'><B>บันทึกข้อตกลงความร่วมมือ</B></div>
    <div class='t-22 text-center'><B>โครงการ {result.ProjectTitle}</B></div>
    <div class='t-16 text-center'><B>ระหว่าง</B></div>
    <div class='t-22 text-center'><B>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B></div>
    <div class='t-22 text-center'><B>กับ</B></div>
    <div class='t-18 text-center'><B>{result.OrgCommonName ?? ""}</B></div>
    <br/>
    <div class='t-16 tab3'>บันทึกข้อตกลงความร่วมมือฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เมื่อ {CommonDAO.ToThaiDateStringCovert(result.Sign_Date ?? DateTime.Now)} ระหว่าง</div>
    <div class='t-16 tab3'>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม โดย {result.OrgCommonName} สำนักงานตั้งอยู่เลขที่ 21 อาคารทีเอสที ทาวเวอร์ ชั้น G,17-18,23 ถนนวิภาวดีรังสิต แขวงจอมพล เขตจตุจักร กรุงเทพมหานคร 10900 ซึ่งต่อไป ในสัญญาฉบับนี้จะเรียกว่า “สสว.” ฝ่ายหนึ่ง กับ</div>
    <div class='t-16 tab3'>“ชื่อเต็มของหน่วยงาน” โดย {result.Requestor} ตำแหน่ง {result.RequestorPosition} ผู้มีอำนาจกระทำการแทนปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ ฉบับลง {CommonDAO.ToThaiDateStringCovert(result.Sign_Date ?? DateTime.Now)} สำนักงานตั้งอยู่เลขที่ ซึ่งต่อไปในสัญญาฉบับนี้จะเรียกว่า “  ” อีกฝ่ายหนึ่ง</div>
    <div class='t-16 tab3'>วัตถุประสงค์ของความร่วมมือ</div>
    <div class='t-16 tab3'>ทั้งสองฝ่ายมีความประสงค์ที่จะร่วมมือกันเพื่อดำเนินการภายใต้โครงการ (ชื่อโครงการที่ระบุไว้ข้างต้น) ซึ่งในบันทึกข้อตกลงฉบับนี้ต่อไปจะเรียกว่า “โครงการ” โดยมีรายละเอียดโครงการแผนการดำเนินงาน แผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) และบรรดาเอกสารแนบท้ายบันทึกข้อตกลงฉบับนี้ ซึ่งให้ถือเป็นส่วนหนึ่งของบันทึกข้อตกลงฉบับนี้ มีระยะเวลา ตั้งแต่วันที่ {CommonDAO.ToThaiDateStringCovert(result.Start_Date ?? DateTime.Now)} จนถึงวันที่ {CommonDAO.ToThaiDateStringCovert(result.End_Date ?? DateTime.Now)} โดยมีวัตถุประสงค์ในการดำเนินโครงการ ดังนี้</div>
    <ul>
        {string.Join("", purposeList.Select((p, i) => $"<li class='t-16 tab3'>{p.Detail}</li>"))}
    </ul>
    <div class='t-16 tab3'><b>ข้อ 1 ขอบเขตความร่วมมือของ “สสว.”</b></div>
    <div class='t-16 tab4'>1.1 ตกลงร่วมดำเนินการโครงการโดยสนับสนุนงบประมาณ จำนวน {result.Contract_Value} บาท ( {CommonDAO.NumberToThaiText(result.Contract_Value ?? 0)} ) ซึ่งได้รวมภาษีมูลค่าเพิ่ม ตลอดจนค่าภาษีอากรอื่น ๆ แล้วให้กับ “ชื่อหน่วยร่วม” และการใช้จ่ายเงินให้เป็นไปตามแผนการจ่ายเงินตามเอกสารแนบท้ายบันทึกข้อตกลงฉบับนี้</div>
    <div class='t-16 tab4'>1.2 ประสานการดำเนินโครงการ เพื่อให้บรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์</div>
    <div class='t-16 tab4'>1.3 กำกับ ติดตามและประเมินผลการดำเนินงานของโครงการ</div>
    <div class='t-16 tab3'><b>ข้อ 2 ขอบเขตความร่วมมือของ “ชื่อหน่วยร่วม”</b></div>
    <div class='t-16 tab4'>2.1 ตกลงที่จะร่วมดำเนินการโครงการตามวัตถุประสงค์ของการโครงการและขอบเขตการดำเนินการตามรายละเอียดโครงการ แผนการดำเนินการ และแผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) ที่แนบท้ายบันทึกข้อตกลงฉบับนี้</div>
    <div class='t-16 tab4'>2.2 ต้องดำเนินโครงการ ปฏิบัติตามแผนการดำเนินงาน แผนการใช้จ่ายเงิน (หรืออาจมีคู่มือการดำเนินโครงการก็ได้) อย่างเคร่งครัดและให้แล้วเสร็จภายในระยะเวลาโครงการ</div>
    <div class='t-16 tab4'>2.3 ต้องประสานการดำเนินโครงการ เพื่อให้โครงการบรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์</div>
    <div class='t-16 tab4'>2.4 ต้องให้ความร่วมมือกับ สสว. ในการกำกับ ติดตามและประเมินผลการดำเนินงานของโครงการ</div>
    <div class='t-16 tab3'><b>ข้อ 3 อื่น ๆ</b></div>
    <div class='t-16 tab4'>3.1 หากฝ่ายใดฝ่ายหนึ่งประสงค์จะขอแก้ไข เปลี่ยนแปลง ขยายระยะเวลาของโครงการ จะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษร และต้องได้รับความยินยอมเป็นลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และต้องทำบันทึกข้อตกลงแก้ไข เปลี่ยนแปลง ขยายระยะเวลา เพื่อลงนามยินยอมทั้งสองฝ่าย</div>
    <div class='t-16 tab4'>3.2 หากฝ่ายใดฝ่ายหนึ่งประสงค์จะขอบอกเลิกบันทึกข้อตกลงความร่วมมือก่อนครบกำหนดระยะเวลาดำเนินโครงการจะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษรไม่น้อยกว่า 30 วัน และต้องได้รับความยินยอมเป็นลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และ “ชื่อหน่วยร่วม” จะต้องคืนเงินในส่วนที่ยังไม่ได้ใช้จ่ายหรือส่วนที่เหลือทั้งหมดพร้อมดอกผล (ถ้ามี) ให้แก่ สสว. ภายใน 15 วัน นับจากวันที่ได้รับหนังสือของฝ่ายที่ยินยอมให้บอกเลิก</div>
    <div class='t-16 tab4'>3.3 สสว. อาจบอกเลิกบันทึกข้อตกลงความร่วมมือได้ทันที หากตรวจสอบ หรือปรากฏข้อเท็จจริงว่า การใช้จ่ายเงินของ “ชื่อหน่วยร่วม” ไม่เป็นไปตามวัตถุประสงค์ของโครงการ แผนการดำเนินงาน และแผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) ทั้งมีสิทธิเรียกเงินคงเหลือพร้อมดอกผล (ถ้ามี) คืนทั้งหมดได้ทันที</div>
    <div class='t-16 tab4'>3.4 ทรัพย์สินใด ๆ และ/หรือ สิทธิใด ๆ ที่ได้มาจากเงินสนับสนุนตามบันทึกข้อตกลงฉบับนี้ เมื่อสิ้นสุดโครงการให้ตกได้แก่ สสว. ทั้งสิ้น เว้นแต่ สสว. จะกำหนดให้เป็นอย่างอื่น</div>
    <div class='t-16 tab4'>3.5 “ชื่อหน่วยร่วม” ต้องไม่ดำเนินการในลักษณะการจ้างเหมา กับหน่วยงาน องค์กร หรือบุคคลอื่น ๆ ยกเว้นกรณีการจัดหา จัดจ้าง เป็นกิจกรรมหรือเป็นเรื่อง ๆ</div>
    <div class='t-16 tab4'>3.6 ในกรณีที่การดำเนินการตามบันทึกข้อตกลงฉบับนี้ เกี่ยวข้องกับข้อมูลส่วนบุคคลและการคุ้มครองทรัพย์สินทางปัญญา “ชื่อหน่วยร่วม” จะต้องปฏิบัติตามกฎหมายว่าด้วยการคุ้มครองข้อมูลส่วนบุคคลและการคุ้มครองทรัพย์สินทางปัญญาอย่างเคร่งครัด และหากเกิดความเสียหายหรือมีการฟ้องร้องใด ๆ “ชื่อหน่วยร่วม” จะต้องเป็นผู้รับผิดชอบต่อการละเมิดบทบัญญัติแห่งกฎหมายดังกล่าวแต่เพียงฝ่ายเดียวโดยสิ้นเชิง</div>
    <div class='t-16 tab4'>บันทึกข้อตกลงความร่วมมือฉบับนี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน ทั้งสองฝ่ายได้อ่านและเข้าใจข้อความโดยละเอียดแล้ว จึงได้ลงลายมือชื่อพร้อมประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยานและยึดถือไว้ฝ่ายละฉบับ</div>
    <br/>
    <div class='center signature'>(ลงชื่อ)....................................................</div>
    <div class='center signature'>({result.OSMEP_Signer ?? ""})</div>
    <div class='center signature'>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</div>
    <br/>
    <div class='center signature'>(ลงชื่อ)....................................................</div>
    <div class='center signature'>({result.OSMEP_Witness ?? ""})</div>
    <div class='center signature'>ชื่อเต็มหน่วยงาน</div>
    <br/>
    <div class='center signature'>(ลงชื่อ)....................................................พยาน</div>
    <div class='center signature'>({result.Contract_Signer ?? ""})</div>
    <br/>
    <div class='center signature'>(ลงชื่อ)....................................................พยาน</div>
    <div class='center signature'>({result.Contract_Witness ?? ""})</div>
</body>
</html>
";

        // You need to inject IConverter _pdfConverter in the constructor for PDF generation
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
    #endregion
}
