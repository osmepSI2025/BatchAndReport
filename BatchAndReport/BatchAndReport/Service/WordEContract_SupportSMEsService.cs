using BatchAndReport.DAO;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Signatures;
using System.Globalization;
using System.Text;
public class WordEContract_SupportSMEsService
{
    private readonly WordServiceSetting _w;
    private readonly Econtract_Report_GADAO _e;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    private readonly E_ContractReportDAO _eContractReportDAO;
    public WordEContract_SupportSMEsService(WordServiceSetting ws, Econtract_Report_GADAO e
        , IConverter pdfConverter, E_ContractReportDAO eContractReportDAO
        )
    {
        _w = ws;
        _e = e;
        _pdfConverter = pdfConverter; // กำหนดค่า DI สำหรับ PDF Converter
        _eContractReportDAO = eContractReportDAO; // กำหนดค่า DI สำหรับ E_ContractReportDAO
    }
    #region  4.1.1.2.2.สัญญารับเงินอุดหนุน

    public async Task<byte[]> OnGetWordContact_SupportSMEsService(string id)
    {
        try
        {
            var result = await _e.GetGAAsync(id);
            if (result == null)
            {
                throw new Exception("ไม่พบข้อมูลสัญญา��ับเงินอุดหนุนสำหรับ SMEs ที่ระบุ");
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

                    // --- 2. Titles ---
                    body.AppendChild(WordServiceSetting.EmptyParagraph());
                    body.AppendChild(WordServiceSetting.CenteredBoldParagraph("สัญญารับเงินอุดหนุน", "44"));
                    body.AppendChild(WordServiceSetting.CenteredBoldParagraph("เพื่อสนับสนุนและยกระดับศักยภาพผู้ประกอบการวิสาหกิจขนาดกลางและขนาดย่อม", "44"));
                    body.AppendChild(WordServiceSetting.CenteredBoldParagraph("ผ่านระบบผู้ให้บริการทางธุรกิจ ปี " + DateTime.Now.Year + " ", "44"));
                    body.AppendChild(WordServiceSetting.RightParagraph("ทะเบียนผู้รับเงินอุดหนุนเลขที่ " + result.TaxID + ""));
                    body.AppendChild(WordServiceSetting.RightParagraph("เลขที่สัญญา " + result.Contract_Number + ""));

                    string signDate = result.ContractSignDate.HasValue ? result.ContractSignDate.Value.ToString("dd MMMM yyyy", new CultureInfo("th-TH")) : "____";
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สัญญาฉบับนี้ทำขึ้น ณ  " + result.SignAddress + "  เมื่อ" + signDate + " ระหว่าง"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สำานักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม โดย นายวชิระ แก้วกอ ผู้มีอำนาจกระทำการ แทนสำนักงานฯ ซึ่งต่อไปในสัญญานี้เรียกว่า “ผู้ให้เงินอุดหนุน” ฝ่ายหนึ่ง กับ"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ผู้ประกอบการวิสาหกิจขนาดกลางและขนาดย่อม ราย " + result.RegType + " )  เลขประจำตัวผู้เสียภาษี " + result.TaxID + " " +
                        "ตั้งอยู่เลขที่ " + result.HQLocationAddressNo + "" +
                        "จังหวัด " + result.HQLocationProvince + " โดย " + result.Contract_Number + "ณ " + result.HQLocationAddressNo + " " + result.HQLocationDistrict + "" + "" +
                        "ตำบล/แขวง" + result.HQLocationDistrict + "อำเภอ/เขต " + result.HQLocationDistrict + "" +
                        "มีสำนักงานใหญ่ . ไปรษณีย์อิเล็กทรอนิกส์ " + result.RegEmail + "" +
                        "บัตรประจำตัวประชาชนเลขที่ " + result.RegIdenID + "" +
                        "ผู้มีอำนาจลงนามผูกพัน (" + result.RegType + " ) ปรากฏตามสำเนา " +
                        "หนังสือรับรอง (นิติบุคคล/ทะเบียนพาณิชย์/วิสาหกิจชุมชน/" +
                        "หุ้นส่วนบริษัท " + result.ContractPartyName + " ลง" + signDate + ") " +
                        "ของสำนักงานทะเบียน ซึ่งต่อไปในสัญญานี้ เรียกว่า “ผู้รับเงินอุดหนุน” อีกฝ่ายหนึ่ง"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ทั้งสองฝ่ายได้ตกลงทำสัญญากัน มีข้อความดังต่อไปนี้"));

                    string stringGrantAmount = CommonDAO.NumberToThaiText(result.GrantAmount ?? 0);
                    string stringGrantStartDate = CommonDAO.ToThaiDateStringCovert(result.GrantStartDate ?? DateTime.Now);
                    string stringGrantEndDate = CommonDAO.ToThaiDateStringCovert(result.GrantEndDate ?? DateTime.Now);
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 1  ผู้ให้เงินอุดหนุนตกลงให้เงินอุดหนุนและผู้รับเงินอุดหนุนตกลงรับเงินอุดหนุน  จำนวน " + result.GrantAmount + " บาท (" + stringGrantAmount + ") ตั้งแต่ " + stringGrantStartDate + "ถึงวันที่ " + stringGrantEndDate + " โดยให้ผู้รับการอุดหนุนเข้ารับการพัฒนา เพื่อใช้จ่ายในการ " + result.SpendingPurpose + " จากการให้ความช่วยเหลือ อุดหนุน จากเงินกองทุนส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม" +
                        " ผ่านผู้ให้บริการ ทางธุรกิจ ปี ๒๕๖๗ ภายใต้โครงการส่งเสริมผู้ประกอบการผ่านระบบ BDS ระยะเวลาดำเนินการ 2 ปี (ปี ๒๕๖๗-๒๕๖๘)  ตามข้อเสนอการพัฒนาซึ่งได้รับอนุมัติจากผู้ให้เงินอุดหนุน ตามระเบียบคณะกรรมการบริหารสำนักงานส่งเสริม วิสาหกิจขนาดกลางและขนาดย่อม ว่าด้วยหลักเกณฑ์ เงื่อนไข และวิธีการให้ความช่วยเหลือ อุดหนุน วิสาหกิจ- ๒ - ขนาดกลางและขนาดย่อม จากเงินกองทุนส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม พ.ศ. 2564 ประกาศ " +
                        "สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เรื่อง เชิญชวนหน่วยงานที่ประสงค์ขึ้นทะเบียนผู้ให้บริการ ทางธุรกิจ เพื่อสนับสนุนและยกระดับศักยภาพผู้ประกอบการวิสาหกิจขนาดกลางและขนาดย่อม และเชิญชวน วิสาหกิจขนาดกลางและขนาดย่อม ยื่นความประสงค์ขอรับความช่วยเหลือ อุดหนุน " +
                        "จากเงินกองทุนส่งเสริม วิสาหกิจขนาดกลางและขนาดย่อม ผ่านผู้ให้บริการทางธุรกิจ ปี ๒๕๖๗ และประกาศสำนักงานส่งเสริมวิสาหกิจ ขนาดกลางและขนาดย่อม เรื่อง เชิญชวนหน่วยงานที่ประสงค์ขึ้นทะเบียนผู้ให้บริการทางธุรกิจ เพื่อสนับสนุน และยกระดับศักยภาพผู้ประกอบการวิสาหกิจขนาดกลางและขนาดย่อม " +
                        "และเชิญชวนวิสาหกิจขนาดกลางและ ขนาดย่อม ยื่นความประสงค์ขอรับความช่วยเหลือ อุดหนุนฯ (ฉบับที่ 2) และผู้รับเงินอุดหนุนต้องดำเนิน กิจกรรมและใช้จ่ายเงินตามแผนการดำเนินงานและแผนการใช้จ่ายที่ระบุไว้ในข้อเสนอการพัฒนาที่ได้รับอนุมัติ อย่างเคร่งครัด และให้ถือว่าเป็นส่วนหนึ่งของสัญญาฉบับน"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 2  ผู้รับเงินอุดหนุนจะต้องสำรองเงินจ่ายไปก่อน แล้วจึงนำต้นฉบับใบเสร็จรับเงินมาเบิกกับ ผู้ให้เงินอุดหนุน วงเงินไม่เกินตามข้อ 1 ทั้งนี้ ผู้ให้เงินอุดหนุนจะสนับสนุนจำนวนเงินตามจำนวนที่จ่ายจริงและ เป็นไปตามสัดส่วนการร่วมค่าใช้จ่ายในการสนับสนุนระหว่างผู้ให้เงินอุดหนุนและผู้รับเงินอุดหนุน โดยสัดส่วน งบประมาณที่ให้การอุดหนุนดังกล่าวต้องเป็นไปตามการจัดกลุ่มและสัดส่วนของผู้ประกอบการ ตามประกาศ แนบท้ายสัญญา"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ในการให้ความช่วยเหลือ อุดหนุน วิสาหกิจขนาดกลางและขนาดย่อม ผ่านผู้ให้บริการทางธุรกิจ    ผู้รับเงินอุดหนุนจะได้รับความช่วยเหลือ อุดหนุน ในโครงการนี้ หรือโครงการให้ความช่วยเหลือ อุดหนุน  ผ่านผู้ให้บริการทางธุรกิจในปีอื่น ๆ ในวงเงินรวมกันสูงสุดไม่เกิน 500,000 บาท (ห้าแสนบาทถ้วน) ตลอดระยะเวลา การดำเนินธุรกิจ  ดังนั้น วงเงินที่ได้รับการอุดหนุนตามสัญญานี้ จะต้องถูกหักจากวงเงินรวมที่ได้รับสิทธิ์ "));


                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 3  เมื่อผู้รับเงินอุดหนุนดำเนินกิจกรรมเข้ารับการพัฒนาเสร็จสมบูรณ์แล้วตามแผนการดำเนิน กิจกรรมในข้อเสนอการพัฒนา และนำส่งรายงานผลการพัฒนาและรายละเอียดที่เกี่ยวข้องมายังผู้ให้เงิน อุดหนุน โดยผู้รับเงินอุดหนุนต้องเบิกค่าใช้จ่ายทันทีหลังจากได้รับการพัฒนาหรือก่อนสิ้นสุดสัญญาฉบับนี้ ภายใน ๓๐ (สามสิบ) วันทำการ นับจากวันที่สิ้นสุดสัญญา"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 4  ผู้รับเงินอุดหนุนยินยอมรับผิดชอบค่าใช้จ่ายส่วนเกินจากการสนับสนุนตามการให้ความ ช่วยเหลือในโครงการนี้ที่ได้กำหนดไว้ รวมทั้งรับผิดชอบภาษีมูลค่าเพิ่ม และภาษีอื่น ๆ (ถ้ามี) ที่เกิดจาก ค่าใช้จ่ายที่ขอรับการอุดหนุน"));


                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 5  เงินที่ผู้รับเงินอุดหนุนได้รับจากโครงการนี้ เป็นเงินที่รวมภาษี และค่าธรรมเนียมต่าง ๆ  ไว้ทั้งหมดแล้ว และถือเป็นรายได้ของวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งจะต้องถูกหักภาษี ณ ที่จ่าย และ ต้องเสียภาษีตามที่กฎหมายกำหนด  และหากวิสาหกิจขนาดกลางและขนาดย่อมเป็นผู้ซึ่งจดทะเบียน ภาษีมูลค่าเพิ่ม จะต้องมีการแสดงรายการคำนวณภาษีมูลค่าเพิ่มไว้ให้ชัดเจนปรากฏไว้ในใบสำคัญการรับเงิน หรือใบเสร็จรับเงิน หรือใบกำกับภาษี ที่ยื่นให้ผู้ให้เงินอุดหนุน โดยวิสาหกิจขนาดกลางและขนาดย่อมมีหน้าที่ จะต้องนำเงินที่ได้รับดังกล่าว ไปประกอบการคำนวณรายได้เพื่อเสียภาษีเงินได้ในปีที่เกิดรายได้ด้วย "));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 6  กรณีการโอนเงินให้แก่ผู้รับเงินอุดหนุน ผู้ให้เงินอุดหนุนจะใช้วิธีการโอนเงินผ่านระบบ อิเล็กทรอนิกส์ และหากมีค่าธรรมเนียมการโอนเงิน ผู้รับเงินอุดหนุนจะเป็นผู้รับผิดชอบค่าธรรมเนียมในการ โอนเงินดังกล่าว"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 7  ผู้รับเงินอุดหนุนจะเปลี่ยนแปลงข้อเสนอการพัฒนาและวงเงินอุดหนุนตามที่ได้รับอนุมัติ จากผู้ให้เงินอุดหนุนได้ ต่อเมื่อผู้รับเงินอุดหนุนได้แจ้งเป็นหนังสือให้ผู้ให้เงินอุดหนุนทราบ และได้รับความ เห็นชอบเป็นหนังสือจากผู้ให้เงินอุดหนุนก่อนทุกครั้ง โดยผู้รับเงินอุดหนุนจะต้องดำเนินการก่อนวันสิ้นสุด สัญญาไม่น้อยกว่า ๓๐ (สามสิบ) วันทำการ"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 8  ผู้รับเงินอุดหนุนจะต้องใช้จ่ายเงินอุดหนุนเพื่อดำเนินการตามข้อเสนอการพัฒนา ซึ่งได้รับการอนุมัติ ให้เป็นไปตามวัตถุประสงค์และกิจกรรมตามข้อเสนอการพัฒนาเท่านั้น โดยผู้รับเงินอุดหนุน ตกลงยินยอมให้ผู้ให้เงินอุดหนุนตรวจสอบผลการปฏิบัติงาน และการใช้จ่ายเงินอุดหนุนที่ได้รับ และผู้รับเงิน อุดหนุนมีหน้าที่ต้องรายงานผลการปฏิบัติงานและการใช้จ่ายเงินอุดหนุนที่รับตามแบบและภายในเวลาที่ กำหนด "));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 9  กรณีที่มีการตรวจพบในภายหลังว่าผู้รับเงินอุดหนุนขาดคุณสมบัติในการรับเงินอุดหนุน ผู้ให้เงินอุดหนุนมีสิทธิบอกเลิกสัญญาได้ทันที หรือในกรณีผู้รับเงินอุดหนุนนำเงินไปใช้ผิดจากวัตถุประสงค์ตาม ข้อเสนอการพัฒนา ผู้รับเงินอุดหนุนจะต้องรับผิดชอบชดใช้เงินอุดหนุนที่ได้รับไปทั้งหมดคืนให้แก่ผู้ให้เงินอุดหนุน ภายใน ๓๐ (สามสิบ) วัน นับแต่วันที่ได้รับหนังสือแจ้งจากผู้ให้เงินอุดหนุน พร้อมด้วยดอกเบี้ยในอัตรา ร้อยละ ๕ (ห้า) ต่อปี นับแต่วันที่ได้รับเงินอุดหนุนจนกว่าจะชดใช้เงินคืนจนครบถ้วนเสร็จสิ้น "));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 10  ในกรณีผู้รับเงินอุดหนุนไม่ปฏิบัติตามสัญญาข้อหนึ่งข้อใด ผู้ให้เงินอุดหนุนจะมีหนังสือแจ้ง ให้ผู้รับเงินอุดหนุนทราบ โดยจะกำหนดระยะเวลาพอสมควรเพื่อให้ปฏิบัติให้ถูกต้องตามสัญญา และหาก ผู้รับเงินอุดหนุนไม่ปฏิบัติภายในระยะเวลาที่กำหนดดังกล่าว ผู้ให้เงินอุดหนุนมีสิทธิบอกเลิกสัญญาได้ทันที โดย มีหนังสือบอกเลิกสัญญาแจ้งให้ผู้รับเงินอุดหนุนทราบ"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 11  ในกรณีที่มีการบอกเลิกสัญญาตามข้อ ๑๐ ผู้รับเงินอุดหนุนจะต้องชดใช้เงินอุดหนุนคืน ให้แก่ผู้ให้เงินอุดหนุนตามจำนวนเงินที่ได้รับทั้งหมด หรือตามจำนวนเงินคงเหลือในวันบอกเลิกสัญญา หรือตาม จำนวนเงินที่ผู้ให้เงินอุดหนุนจะพิจารณาตามความเหมาะสมแล้วแต่กรณี ซึ่งผู้ให้เงินอุดหนุนจะแจ้งเป็นหนังสือ พร้อมการบอกเลิกสัญญา ให้ผู้รับเงินอุดหนุนทราบว่าต้องชดใช้เงินคืนจำนวนเท่าใด โดยผู้รับเงินอุดหนุนต้อง ชำระเงินดังกล่าวพร้อมด้วยดอกเบี้ยในอัตรา ร้อยละ ๕ (ห้า) ต่อปี นับแต่วันบอกเลิกสัญญาจนถึงวันที่ชดใช้ เงินคืนจนครบถ้วนเสร็จสิ้น ทั้งนี้ ในกรณีเกิดความเสียหายอย่างหนึ่งอย่างใดแก่ผู้ให้เงินอุดหนุน ผู้ให้เงิน อุดหนุนมีสิทธิที่จะเรียกค่าเสียหายจากผู้รับเงินอุดหนุนอีกด้วย"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 12  ผู้รับเงินอุดหนุนต้องปฏิบัติตามเงื่อนไขที่กำหนดไว้ในระเบียบและประกาศแนบท้าย สัญญานี้"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 13  ที่อยู่ของผู้รับเงินอุดหนุนที่ปรากฏในสัญญานี้ ให้ถือว่าเป็นภูมิลำเนาของผู้รับเงินอุดหนุน การส่งหนังสือ คำบอกกล่าวทวงถาม จดหมาย หรือเอกสารอื่นใดไปยังผู้รับเงินอุดหนุน ให้ส่งไปยังภูมิลำเนา   ผู้รับเงินอุดหนุนดังกล่าว และให้ถือว่าเป็นการส่งโดยชอบ โดยถือว่าผู้รับเงินอุดหนุนได้ทราบข้อความ ในเอกสารดังกล่าวนับแต่วันที่หนังสือ คำบอกกล่าวทวงถาม จดหมาย หรือเอกสารอื่นใดไปถึงภูมิลำเนา ของผู้รับเงินอุดหนุน ไม่ว่าผู้รับเงินอุดหนุนหรือบุคคลอื่นใดที่พักอาศัยอยู่ในภูมิลำเนาของผู้รับเงินอุดหนุนจะ ได้รับหนังสือ คำบอกกล่าวทวงถาม จดหมาย หรือเอกสารนั้นไว้หรือไม่ก็ตาม"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ถ้าผู้รับเงินอุดหนุนเปลี่ยนแปลงสถานที่อยู่ หรือไปรษณีย์อิเล็กทรอนิกส์ (E-mail) ผู้รับเงินอุดหนุน มีหน้าที่แจ้งให้ผู้ให้เงินอุดหนุนทราบภายใน ๗ (เจ็ด) วัน นับแต่วันเปลี่ยนแปลงสถานที่อยู่หรือไปรษณีย์ อิเล็กทรอนิกส์ (E-mail) หากผู้รับเงินอุดหนุนไม่แจ้งการเปลี่ยนแปลงสถานที่อยู่และผู้ให้เงินอุดหนุนได้ส่ง หนังสือ คำบอกกล่าวทวงถาม จดหมาย หรือเอกสารอื่นใดไปยังผู้รับเงินอุดหนุนตามที่อยู่ที่ปรากฏในสัญญานี้ ให้ถือว่าผู้รับเงินอุดหนุนได้ทราบข้อความในเอกสารดังกล่าวโดยชอบตามวรรคหนึ่งแล้ว"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สัญญานี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน คู่สัญญาได้อ่านและเข้าใจข้อความ โดยละเอียดตลอดแล้ว และได้ตกลงกันให้ถือว่าได้ส่ง ณ ที่ทำการงานของผู้ให้เงินอุดหนุน หรือได้รับ ณ ที่ทำการ งานของผู้รับเงินอุดหนุน จึงได้ลงลายมือชื่อ พร้อมทั้งประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยาน และคู่สัญญา ต่างยึดถือไว้ ฝ่ายละหนึ่งฉบับ"));




                    // --- Signature Table: 2 columns for each row ---
                    var signatureTable = new Table(
                        new TableProperties(
                            new TableWidth { Width = "10000", Type = TableWidthUnitValues.Pct },
                            new TableBorders(
                                new TopBorder { Val = BorderValues.None },
                                new BottomBorder { Val = BorderValues.None },
                                new LeftBorder { Val = BorderValues.None },
                                new RightBorder { Val = BorderValues.None },
                                new InsideHorizontalBorder { Val = BorderValues.None },
                                new InsideVerticalBorder { Val = BorderValues.None }
                            )
                        ),
                        // Row 1: Signatures
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                                WordServiceSetting.NormalParagraph("(ลงชื่อ)....................................................ผู้ให้เงินอุดหนุน", JustificationValues.Center, "32")
                            ),
                            new TableCell(
                                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                                WordServiceSetting.NormalParagraph("(ลงชื่อ)....................................................ผู้รับเงินอุดหนุน", JustificationValues.Center, "32")
                            )
                        ),
                        // Row 2: Names
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                                WordServiceSetting.NormalParagraph("(นายวชิระ แก้วกอ)", JustificationValues.Center, "32"),
                                WordServiceSetting.NormalParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม", JustificationValues.Center, "32")
                            ),
                            new TableCell(
                                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                                WordServiceSetting.NormalParagraph("(....................................................)", JustificationValues.Center, "32"),
                                WordServiceSetting.NormalParagraph("ผู้ประกอบการวิสาหกิจขนาดกลางและขนาดย่อม", JustificationValues.Center, "32"),
                                WordServiceSetting.NormalParagraph("ราย....................................................", JustificationValues.Center, "32")
                            )
                        ),
                        // Row 3: Witnesses
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                                WordServiceSetting.NormalParagraph("(ลงชื่อ)....................................................พยาน", JustificationValues.Center, "32"),
                                WordServiceSetting.NormalParagraph("(นางสาวนิธิวดี สมบูรณ์)", JustificationValues.Center, "32")
                            ),
                            new TableCell(
                                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                                WordServiceSetting.NormalParagraph("(ลงชื่อ)....................................................พยาน", JustificationValues.Center, "32"),
                                WordServiceSetting.NormalParagraph("(....................................................)", JustificationValues.Center, "32")
                            )
                        ),
                        // Row 4: Additional Witness
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "100" }),
                                WordServiceSetting.NormalParagraph("(ลงชื่อ)....................................................พยาน", JustificationValues.Center, "32"),
                                WordServiceSetting.NormalParagraph("(นางสาวพัชณีภานต์ นาคบัว)", JustificationValues.Center, "32")
                            )
                        )
                    );
                    body.AppendChild(signatureTable);


                    WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);
                }
                stream.Position = 0;
                return stream.ToArray();
            }
        }
        catch (Exception ex)
        {
            // Log the exception or handle it as needed
            throw new Exception("Error generating Word contract for Support SMEs: " + ex.Message, ex);
        }

    }
    #endregion  4.1.1.2.2.สัญญารับเงินอุดหนุน

    public async Task<string> OnGetWordContact_SupportSMEsService_HtmlToPDF(string id,string typeContact)
    {
        var result = await _e.GetGAAsync(id);
   
        if (result == null)
            throw new Exception("ไม่พบข้อมูลสัญญารับเงินอุดหนุนสำหรับ SMEs ที่ระบุ");

        // อ่านไฟล์โลโก้และแปลงเป็น Base64
        var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
        string logoBase64 = "";
        if (System.IO.File.Exists(logoPath))
        {
            var bytes = System.IO.File.ReadAllBytes(logoPath);
            logoBase64 = Convert.ToBase64String(bytes);
        }

        // สร้าง path ฟอนต์แบบ absolute
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");
        var cssPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "css", "contract.css").Replace("\\", "/");

        string signDate = CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now);
        string stringGrantAmount = CommonDAO.NumberToThaiText(result.GrantAmount ?? 0);
        string stringGrantStartDate = CommonDAO.ToThaiDateStringCovert(result.GrantStartDate ?? DateTime.Now);
        string stringGrantEndDate = CommonDAO.ToThaiDateStringCovert(result.GrantEndDate ?? DateTime.Now);



        var signlist = await _eContractReportDAO.GetSignNameAsync(id, typeContact);
        var signatoryHtml = new StringBuilder();
        var companySealHtml = new StringBuilder();

        foreach (var signer in signlist)
        {
            string signatureHtml;
            string companySeal = ""; // Initialize to avoid unassigned variable warning

            // Fix CS8602: Use null-conditional operator for Position and Company_Seal
            if (signer?.Signatory_Type == "CP_S" && !string.IsNullOrEmpty(signer?.Company_Seal))
            {
                try
                {
                    var contentStart = signer.Company_Seal.IndexOf("<content>") + "<content>".Length;
                    var contentEnd = signer.Company_Seal.IndexOf("</content>");
                    var base64 = signer.Company_Seal.Substring(contentStart, contentEnd - contentStart);

                    companySeal = $@"
<div class='t-16 text-center tab1'>
                <img src='data:image/png;base64,{base64}' alt='signature' style='max-height: 80px;' />
            </div>";

                    companySealHtml.AppendLine($@"
    <div class='text-center'>
        {companySeal}      
    </div>
<div class='t-16 text-center tab1'>(ตราประทับ บริษัท)</div>

");
                }
                catch
                {
                    companySeal = "<div class='t-16 text-center tab1'>(ตราประทับ บริษัท)</div>";
                }
            }

            if (!string.IsNullOrEmpty(signer?.DS_FILE) && signer.DS_FILE.Contains("<content>"))
            {
                try
                {
                    var contentStart = signer.DS_FILE.IndexOf("<content>") + "<content>".Length;
                    var contentEnd = signer.DS_FILE.IndexOf("</content>");
                    var base64 = signer.DS_FILE.Substring(contentStart, contentEnd - contentStart);

                    signatureHtml = $@"<div class='t-16 text-center tab1'>
                <img src='data:image/png;base64,{base64}' alt='signature' style='max-height: 80px;' />
            </div>";
                }
                catch
                {
                    signatureHtml = "<div class='t-16 text-center tab1'>(ลงชื่อ..........)</div>";
                }
            }
            else
            {
                signatureHtml = "<div class='t-16 text-center tab1'>(ลงชื่อ..........)</div>";
            }

            signatoryHtml.AppendLine($@"
    <div class='sign-single-right'>
        {signatureHtml}
        <div class='t-16 text-center tab1'>({signer?.Signatory_Name})</div>
        <div class='t-16 text-center tab1'>{signer?.BU_UNIT}</div>
    </div>");

            signatoryHtml.Append(companySealHtml);
        }


        // สร้าง HTML สำหรับสัญญา 


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
           .tab1 {{ text-indent: 48px; text-align: justify;  }}
        .tab2 {{ text-indent: 96px;  text-align: left; }}
        .tab3 {{ text-indent: 144px; text-align: left; }}
        .tab4 {{ text-indent: 192px;  text-align: left;}}
       .normal {{text-align: justify;
        text-align-last: justify;
        width: 100%;
        display: block;
        min-width: 100%;
  letter-spacing: 0.1em; /* เพิ่มช่องไฟเล็กน้อย */
    }}
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
        p {{
            margin: 0;
            padding: 0;
        }}
    </style>
</head>
<body>

    <div class='text-center'>
         <img src='data:image/jpeg;base64,{logoBase64}' width='240' height='80' />
    </div>
</br>
</br>
    <div class='t-22 text-center'><B>สัญญารับเงินอุดหนุน</B></div>
    <div class='t-22 text-center'><B>เพื่อสนับสนุนและยกระดับศักยภาพผู้ประกอบการวิสาหกิจขนาดกลางและขนาดย่อม</B></div>
    <div class='t-22 text-center'><B>ผ่านระบบผู้ให้บริการทางธุรกิจ ปี {DateTime.Now.Year}</B></div>
</br>
    <div class=' t-16 text-right'>ทะเบียนผู้รับเงินอุดหนุนเลขที่ {result.TaxID}</div>
    <div class=' t-16 text-right'>เลขที่สัญญา {result.Contract_Number}</div>
</br>
    <p class='t-16 tab3'>สัญญาฉบับนี้ทำขึ้น ณ {result.SignAddress} เมื่อ {signDate} ระหว่าง</P>
    <p class='t-16 tab3'><B>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</B> โดย {result.SignatoryName} ผู้มีอำนาจกระทำการแทนสำนักงานฯ ซึ่งต่อไปในสัญญานี้เรียกว่า “ผู้ให้เงินอุดหนุน” ฝ่ายหนึ่ง กับ</P>
    <p class='t-16 tab3'><B>ผู้ประกอบการวิสาหกิจขนาดกลางและขนาดย่อม</B> ราย {result.RegType} ซึ่งจดทะเบียนเป็น {result.RegType} เลขประจำตัวผู้เสียภาษี {result.TaxID} 
        ณ {result.HQLocationAddressNo} {result.HQLocationDistrict} มีสำนักงานใหญ่
        ตั้งอยู่เลขที่ {result.HQLocationAddressNo}
        ตำบล/แขวง {result.HQLocationDistrict} อำเภอ/เขต {result.HQLocationDistrict} จังหวัด {result.HQLocationProvince} 
        ไปรษณีย์อิเล็กทรอนิกส์(E-mail) {result.RegEmail} โดย {result.Contract_Number} บัตรประจำตัวประชาชนเลขที่ {result.RegIdenID}
        ผู้มีอำนาจลงนามผูกพัน {result.RegType} ปรากฏตามสำเนา
        หนังสือรับรอง {result.RegType} ของสำนักงานทะเบียน 
        หุ้นส่วนบริษัท {result.ContractPartyName} ลง {signDate})
        ซึ่งต่อไปในสัญญานี้ เรียกว่า “ผู้รับเงินอุดหนุน” อีกฝ่ายหนึ่ง
    </P>
    <p class='t-16 tab3'>ทั้งสองฝ่ายได้ตกลงทำสัญญากัน มีข้อความดังต่อไปนี้</P>
    <p class='t-16 tab3'>ข้อ 1 ผู้ให้เงินอุดหนุนตกลงให้เงินอุดหนุนและผู้รับเงินอุดหนุนตกลงรับเงินอุดหนุน  </p>
<p class='t-16 normal'>จำนวน {result.GrantAmount?.ToString("N2") ?? "0.00"} บาท ({stringGrantAmount})</p>
<p class='t-16 normal'>ตั้งแต่ {stringGrantStartDate} ถึงวันที่ {stringGrantEndDate}</p>
<p class='t-16 normal'>โดยให้ผู้รับการอุดหนุนเข้ารับการพัฒนา เพื่อใช้จ่ายในการ {result.SpendingPurpose}</p>
  <p class='t-16'>จากการให้ความช่วยเหลือ อุดหนุน จากเงินกองทุนส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม
        ผ่านผู้ให้บริการ ทางธุรกิจ ปี ๒๕๖๗ ภายใต้โครงการส่งเสริมผู้ประกอบการผ่านระบบ BDS ระยะเวลาดำเนินการ ๒ ปี 
(ปี ๒๕๖๗-๒๕๖๘) ตามข้อเสนอการพัฒนาซึ่งได้รับอนุมัติจากผู้ให้เงินอุดหนุน ตามระเบียบคณะกรรมการ
บริหารสำนักงานส่งเสริม วิสาหกิจขนาดกลางและขนาดย่อม ว่าด้วยหลักเกณฑ์ เงื่อนไข และวิธีการให้ความ
ช่วยเหลือ อุดหนุน วิสาหกิจ- 2 - ขนาดกลางและขนาดย่อม จากเงินกองทุนส่งเสริมวิสาหกิจขนาดกลางและ
ขนาดย่อม พ.ศ. ๒๕๖๔ ประกาศ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เรื่อง เชิญชวน
หน่วยงานที่ประสงค์ขึ้นทะเบียนผู้ให้บริการ ทางธุรกิจ เพื่อสนับสนุนและยกระดับศักยภาพผู้ประกอบการ
วิสาหกิจขนาดกลางและขนาดย่อม และเชิญชวน วิสาหกิจขนาดกลางและขนาดย่อม ยื่นความประสงค์ขอรับ
ความช่วยเหลือ อุดหนุนจากเงินกองทุนส่งเสริม วิสาหกิจขนาดกลางและขนาดย่อม ผ่านผู้ให้บริการทางธุรกิจ 
ปี ๒๕๖๗ และประกาศสำนักงานส่งเสริมวิสาหกิจ ขนาดกลางและขนาดย่อม เรื่อง เชิญชวนหน่วยงานที่
ประสงค์ขึ้นทะเบียนผู้ให้บริการทางธุรกิจ เพื่อสนับสนุน และยกระดับศักยภาพผู้ประกอบการวิสาหกิจขนาด
กลางและขนาดย่อมและเชิญชวนวิสาหกิจขนาดกลางและ ขนาดย่อม ยื่นความประสงค์ขอรับความช่วยเหลือ 
อุดหนุนฯ (ฉบับที่ ๒) และผู้รับเงินอุดหนุนต้องดำเนิน กิจกรรมและใช้จ่ายเงินตามแผนการดำเนินงานและ
แผนการใช้จ่ายที่ระบุไว้ในข้อเสนอการพัฒนาที่ได้รับอนุมัติ อย่างเคร่งครัด และให้ถือว่าเป็นส่วนหนึ่งของ
สัญญาฉบับนี้
    </P>
    <p class='t-16 tab3'>ข้อ ๒ ผู้รับเงินอุดหนุนจะต้องสำรองเงินจ่ายไปก่อน แล้วจึงนำต้นฉบับใบเสร็จรับเงินมาเบิก
กับ ผู้ให้เงินอุดหนุน วงเงินไม่เกินตามข้อ ๑ ทั้งนี้ ผู้ให้เงินอุดหนุนจะสนับสนุนจำนวนเงินตามจำนวนที่
จ่ายจริงและ เป็นไปตามสัดส่วนการร่วมค่าใช้จ่ายในการสนับสนุนระหว่างผู้ให้เงินอุดหนุนและผู้รับเงินอุดหนุน โดยสัดส่วน งบประมาณที่ให้การอุดหนุนดังกล่าวต้องเป็นไปตามการจัดกลุ่มและสัดส่วนของผู้ประกอบการ ตามประกาศ แนบท้ายสัญญา</P>
   
<p class='t-16 tab3'>ในการให้ความช่วยเหลือ อุดหนุน วิสาหกิจขนาดกลางและขนาดย่อม ผ่านผู้ให้บริการทาง
ธุรกิจ ผู้รับเงินอุดหนุนจะได้รับความช่วยเหลือ อุดหนุน ในโครงการนี้ หรือโครงการให้ความช่วยเหลือ อุดหนุน ผ่านผู้ให้บริการทางธุรกิจในปีอื่นๆ ในวงเงินรวมกันสูงสุดไม่เกิน ๕๐๐,๐๐๐.๐๐ บาท (ห้าแสนบาทถ้วน) ตลอดระยะ
เวลา การดำเนินธุรกิจ  ดังนั้น วงเงินที่ได้รับการอุดหนุนตามสัญญานี้ จะต้องถูกหักจากวงเงินรวมที่ได้รับสิทธิ์ <br></P>
  
<p class='t-16 tab3'>ข้อ ๓ เมื่อผู้รับเงินอุดหนุนดำเนินกิจกรรมเข้ารับการพัฒนาเสร็จสมบูรณ์แล้วตามแผนการ
ดำเนิน กิจกรรมในข้อเสนอการพัฒนา และนำส่งรายงานผลการพัฒนาและรายละเอียดที่เกี่ยวข้องมายังผู้
ให้เงิน อุดหนุน โดยผู้รับเงินอุดหนุนต้องเบิกค่าใช้จ่ายทันทีหลังจากได้รับการพัฒนาหรือก่อนสิ้นสุดสัญญา
ฉบับนี้ ภายใน ๓๐ (สามสิบ) วันทำการ นับจากวันที่สิ้นสุดสัญญา</P>
    <p class='t-16 tab3'>ข้อ ๔ ผู้รับเงินอุดหนุนยินยอมรับผิดชอบค่าใช้จ่ายส่วนเกินจากการสนับสนุนตามการให้ความ ช่วยเหลือในโครงการนี้ที่ได้กำหนดไว้ รวมทั้งรับผิดชอบภาษีมูลค่าเพิ่ม และภาษีอื่น ๆ (ถ้ามี) ที่เกิดจาก ค่าใช้จ่ายที่ขอรับการอุดหนุน</P>
    <p class='t-16 tab3'>ข้อ ๕ เงินที่ผู้รับเงินอุดหนุนได้รับจากโครงการนี้ เป็นเงินที่รวมภาษี และค่าธรรมเนียมต่างๆ ไว้ทั้งหมดแล้ว และถือเป็นรายได้ของวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งจะต้องถูกหักภาษี ณ ที่จ่าย และ ต้องเสียภาษีตามที่กฎหมายกำหนด และหากวิสาหกิจขนาดกลางและขนาดย่อมเป็นผู้ซึ่งจดทะเบียน ภาษีมูลค่าเพิ่ม จะต้องมีการแสดงรายการคำนวณภาษีมูลค่าเพิ่มไว้ให้ชัดเจนปรากฏไว้ในใบสำคัญการรับเงิน หรือใบเสร็จรับเงิน หรือใบกำกับภาษี ที่ยื่นให้ผู้ให้เงินอุดหนุน โดยวิสาหกิจขนาดกลางและขนาดย่อมมีหน้าที่ จะต้องนำเงินที่ได้รับดังกล่าว ไปประกอบการคำนวณรายได้เพื่อเสียภาษีเงินได้ในปีที่เกิดรายได้ด้วย </P>
    <p class='t-16 tab3'>ข้อ ๖ กรณีการโอนเงินให้แก่ผู้รับเงินอุดหนุน ผู้ให้เงินอุดหนุนจะใช้วิธีการโอนเงินผ่านระบบ อิเล็กทรอนิกส์ และหากมีค่าธรรมเนียมการโอนเงิน ผู้รับเงินอุดหนุนจะเป็นผู้รับผิดชอบค่าธรรมเนียมในการ โอนเงินดังกล่าว</P>
    <p class='t-16 tab3'>ข้อ ๗ ผู้รับเงินอุดหนุนจะเปลี่ยนแปลงข้อเสนอการพัฒนาและวงเงินอุดหนุนตามที่ได้รับ
อนุมัติ จากผู้ให้เงินอุดหนุนได้ ต่อเมื่อผู้รับเงินอุดหนุนได้แจ้งเป็นหนังสือให้ผู้ให้เงินอุดหนุนทราบ และ
ได้รับความ เห็นชอบเป็นหนังสือจากผู้ให้เงินอุดหนุนก่อนทุกครั้ง โดยผู้รับเงินอุดหนุนจะต้องดำเนินการ
ก่อนวันสิ้นสุด สัญญาไม่น้อยกว่า ๓๐ (สามสิบ) วันทำการ</P>
    <p class='t-16 tab3'>ข้อ ๘ ผู้รับเงินอุดหนุนจะต้องใช้จ่ายเงินอุดหนุนเพื่อดำเนินการตามข้อเสนอการพัฒนา ซึ่งได้รับการอนุมัติ ให้เป็นไปตามวัตถุประสงค์และกิจกรรมตามข้อเสนอการพัฒนาเท่านั้น โดยผู้รับเงินอุดหนุน ตกลงยินยอมให้ผู้ให้เงินอุดหนุนตรวจสอบผลการปฏิบัติงาน และการใช้จ่ายเงินอุดหนุนที่ได้รับ และผู้รับเงิน อุดหนุนมีหน้าที่ต้องรายงานผลการปฏิบัติงานและการใช้จ่ายเงินอุดหนุนที่รับตามแบบและภายในเวลาที่ กำหนด </P>

    <p class='t-16 tab3'>ข้อ ๙ กรณีที่มีการตรวจพบในภายหลังว่าผู้รับเงินอุดหนุนขาดคุณสมบัติในการรับเงินอุดหนุน 
ผู้ให้เงินอุดหนุนมีสิทธิบอกเลิกสัญญาได้ทันที หรือในกรณีผู้รับเงินอุดหนุนนำเงินไปใช้ผิดจากวัตถุประสงค์ตาม 
ข้อเสนอการพัฒนา ผู้รับเงินอุดหนุนจะต้องรับผิดชอบชดใช้เงินอุดหนุนที่ได้รับไปทั้งหมดคืนให้แก่ผู้ให้เงินอุด
หนุน ภายใน ๓๐ (สามสิบ) วัน นับแต่วันที่ได้รับหนังสือแจ้งจากผู้ให้เงินอุดหนุน พร้อมด้วยดอกเบี้ยในอัตรา
ร้อยละ 5 (ห้า) ต่อปี นับแต่วันที่ได้รับเงินอุดหนุนจนกว่าจะชดใช้เงินคืนจนครบถ้วนเสร็จสิ้น </P>

    <p class='t-16 tab3'>ข้อ ๑๐ ในกรณีผู้รับเงินอุดหนุนไม่ปฏิบัติตามสัญญาข้อหนึ่งข้อใด ผู้ให้เงินอุดหนุนจะมีหนัง
สือแจ้ง ให้ผู้รับเงินอุดหนุนทราบ โดยจะกำหนดระยะเวลาพอสมควรเพื่อให้ปฏิบัติให้ถูกต้องตามสัญญา 
และหาก ผู้รับเงินอุดหนุนไม่ปฏิบัติภายในระยะเวลาที่กำหนดดังกล่าว ผู้ให้เงินอุดหนุนมีสิทธิบอกเลิกสัญญา
ได้ทันที โดย มีหนังสือบอกเลิกสัญญาแจ้งให้ผู้รับเงินอุดหนุนทราบ</P>
    <p class='t-16 tab3'>ข้อ ๑๑ ในกรณีที่มีการบอกเลิกสัญญาตามข้อ ๑๐ ผู้รับเงินอุดหนุนจะต้องชดใช้เงินอุดหนุน
คืน ให้แก่ผู้ให้เงินอุดหนุนตามจำนวนเงินที่ได้รับทั้งหมด หรือตามจำนวนเงินคงเหลือในวันบอกเลิกสัญญา 
หรือตาม จำนวนเงินที่ผู้ให้เงินอุดหนุนจะพิจารณาตามความเหมาะสมแล้วแต่กรณี ซึ่งผู้ให้เงินอุดหนุนจะแจ้ง
เป็นหนังสือ พร้อมการบอกเลิกสัญญา ให้ผู้รับเงินอุดหนุนทราบว่าต้องชดใช้เงินคืนจำนวนเท่าใด โดยผู้รับเงิน
อุดหนุนต้อง ชำระเงินดังกล่าวพร้อมด้วยดอกเบี้ยในอัตรา ร้อยละ ๕ (ห้า) ต่อปี นับแต่วันบอกเลิกสัญญา
จนถึงวันที่ชดใช้ เงินคืนจนครบถ้วนเสร็จสิ้น ทั้งนี้ ในกรณีเกิดความเสียหายอย่างหนึ่งอย่างใดแก่ผู้ให้เงิน
อุดหนุน ผู้ให้เงิน อุดหนุนมีสิทธิที่จะเรียกค่าเสียหายจากผู้รับเงินอุดหนุนอีกด้วย</P>
    <p class='t-16 tab3'>ข้อ ๑๒ ผู้รับเงินอุดหนุนต้องปฏิบัติตามเงื่อนไขที่กำหนดไว้ในระเบียบและประกาศแนบท้าย สัญญานี้</P>

    <p class='t-16 tab3'>ข้อ ๑๓ ที่อยู่ของผู้รับเงินอุดหนุนที่ปรากฏในสัญญานี้ ให้ถือว่าเป็นภูมิลำเนาของผู้รับเงิน
อุดหนุน การส่งหนังสือ คำบอกกล่าวทวงถาม จดหมาย หรือเอกสารอื่นใดไปยังผู้รับเงินอุดหนุน ให้ส่ง
ไปยังภูมิลำเนา ผู้รับเงินอุดหนุนดังกล่าว และให้ถือว่าเป็นการส่งโดยชอบ โดยถือว่าผู้รับเงินอุดหนุน
ได้ทราบข้อความ ในเอกสารดังกล่าวนับแต่วันที่หนังสือ คำบอกกล่าวทวงถาม จดหมาย หรือเอกสารอื่นใด
ไปถึงภูมิลำเนา ของผู้รับเงินอุดหนุน ไม่ว่าผู้รับเงินอุดหนุนหรือบุคคลอื่นใดที่พักอาศัยอยู่ในภูมิลำเนาของผู้
รับเงินอุดหนุนจะ ได้รับหนังสือ คำบอกกล่าวทวงถาม จดหมาย หรือเอกสารนั้นไว้หรือไม่ก็ตาม</P>
    <p class='t-16 tab3'>ถ้าผู้รับเงินอุดหนุนเปลี่ยนแปลงสถานที่อยู่ หรือไปรษณีย์อิเล็กทรอนิกส์ (E-mail) ผู้รับเงิน
อุดหนุน มีหน้าที่แจ้งให้ผู้ให้เงินอุดหนุนทราบภายใน ๗ (เจ็ด) วัน นับแต่วันเปลี่ยนแปลงสถานที่อยู่หรือ
ไปรษณีย์อิเล็กทรอนิกส์ (E-mail) หากผู้รับเงินอุดหนุนไม่แจ้งการเปลี่ยนแปลงสถานที่อยู่และผู้ให้เงิน
อุดหนุนได้ส่ง หนังสือ คำบอกกล่าวทวงถาม จดหมาย หรือเอกสารอื่นใดไปยังผู้รับเงินอุดหนุนตามที่อยู่ที่
ปรากฏในสัญญานี้ ให้ถือว่าผู้รับเงินอุดหนุนได้ทราบข้อความในเอกสารดังกล่าวโดยชอบตามวรรคหนึ่งแล้ว</P>
    <p class='t-16 tab3'>สัญญานี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน คู่สัญญาได้อ่านและเข้าใจข้อความ โดยละเอียดตลอดแล้ว และได้ตกลงกันให้ถือว่าได้ส่ง ณ ที่ทำการงานของผู้ให้เงินอุดหนุน หรือได้รับ ณ ที่ทำการ งานของผู้รับเงินอุดหนุน จึงได้ลงลายมือชื่อ พร้อมทั้งประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยาน และคู่สัญญา ต่างยึดถือไว้ ฝ่ายละหนึ่งฉบับ</P>

</br>
</br>
{signatoryHtml}
</body>
</html>
";
        //var doc = new HtmlToPdfDocument()
        //{
        //    GlobalSettings = {
        //    PaperSize = PaperKind.A4,
        //    Orientation = Orientation.Portrait,
        //    Margins = new MarginSettings
        //    {
        //        Top = 20,
        //        Bottom = 20,
        //        Left = 20,
        //        Right = 20
        //    }
        //},
        //    Objects = {
        //    new ObjectSettings() {
        //        HtmlContent = html,
        //        FooterSettings = new FooterSettings
        //        {
        //            FontName = "THSarabunNew",
        //            FontSize = 6,
        //            Line = false,
        //            Center = "[page] / [toPage]" // Thai page numbering
        //        }
        //    }
        //}
        //};

        //var pdfBytes = _pdfConverter.Convert(doc);
        return html;
    }
}
