using BatchAndReport.DAO;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Threading.Tasks;


public class WordEContract_ConsultantService
{
    private readonly WordServiceSetting _w;
    private readonly Econtract_Report_CTRDAO _e;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    private readonly EContractDAO _eContractDAO;
    private readonly E_ContractReportDAO _eContractReportDAO;
    public WordEContract_ConsultantService(WordServiceSetting ws
        , Econtract_Report_CTRDAO e
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
    #region 4.1.1.2.14.สัญญาจ้างที่ปรึกษา
    public async Task<byte[]> OnGetWordContact_ConsultantService(string id)
    {
        try
        {
            var result = await _e.GetCTRAsync(id);
            if (result == null)
            {
                throw new Exception("ไม่พบข้อมูลสัญญาจ้างที่ปรึกษา");
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

                    body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("แบบสัญญา", "000000", "36"));
                    body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("สัญญาจ้างผู้เชี่ยวชาญรายบุคคลหรือจ้างบริษัทที่ปรึกษา", "000000", "36"));
                    // 2.Document title and subtitle
                    body.AppendChild(WordServiceSetting.EmptyParagraph());
                    body.AppendChild(WordServiceSetting.RightParagraph("สัญญาเลขที่………….…… (1)...........……..……..."));



                    string datestring = CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now);
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สัญญาฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.)\r\n"
                        + "ตำบล/แขวง ทุ่งสองห้อง อำเภอ/เขต หลักสี่\r\n" +
                    "จังหวัด กรุงเทพ เมื่อ" + datestring + "\r\n" +
                    "ระหว่าง " + result.Contract_Organization + "\r\n" +
                    "โดย " + result.SignatoryName + "\r\n" +
                    "ซึ่งต่อไปในสัญญานี้เรียกว่า “ผู้ว่าจ้าง” ฝ่ายหนึ่ง กับ…" + result.ContractorName + "", null, "32"));

                    // นิติบุคคล
                    if (result.ContractorType == "นิติบุคคล")
                    {
                        body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ซึ่งจดทะเบียนเป็นนิติบุคคล ณ " + result.ContractorName + " มี\r\n" +
                     "สำนักงานใหญ่อยู่เลขที่ " + result.ContractorAddressNo + "ถนน " + result.ContractorStreet + " ตำบล/แขวง " + result.ContractorSubDistrict + "\r\n" +
                     "อำเภอ/เขต " + result.ContractorDistrict + " จังหวัด " + result.ContractorProvince + " \r\nโดย " + result.ContractorSignatoryName + "" +
                     "มีอำนาจลงนามผูกพันนิติบุคคลปรากฏตามหนังสือรับรองของสำนักงานทะเบียนหุ้นส่วนบริษัท ……………\r\n" +
                     "ลงวันที่ " + CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now) + "  (5)(และหนังสือมอบอำนาจลง " + CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now) + ") แนบท้ายสัญญานี้\r\n"
                   , null, "32"));
                    }
                    else
                    {
                        //บุคคลธรรมดา
                        body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(6)(ในกรณีที่ผู้รับจ้างเป็นบุคคลธรรมดาให้ใช้ข้อความว่า กับ " + result.ContractorName + "\r\n" +
                          "อยู่บ้านเลขที่ " + result.ContractorAddressNo + "ถนน " + result.ContractorStreet + " ตำบล/แขวง " + result.ContractorSubDistrict + "\r\n" +
                        "อำเภอ/เขต " + result.ContractorDistrict + " จังหวัด " + result.ContractorProvince + " \r\n" +
                        " ผู้ถือบัตรประจำตัวประชาชนเลขที่ " + result.CitizenId + " ดังปรากฏตามสำเนาบัตรประจำตัวประชาชนแนบท้ายสัญญานี้) ซึ่งต่อไปใน\r\n" +
                        "สัญญานี้เรียกว่า “ที่ปรึกษา” อีกฝ่ายหนึ่ง กับ…" + result.ContractorName + "", null, "32"));

                    }

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("คู่สัญญาได้ตกลงกันมีข้อความดังต่อไปนี้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 1 ข้อตกลงว่าจ้าง", null, "32", true));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("1.1 ผู้ว่าจ้างตกลงจ้างและที่ปรึกษาตกลงรับจ้างปฏิบัติงานตามโครงการ "+result.ProjectName+":" +
                        " "+result.ProjectDesc+" "+result.ProjectReference +"ตามข้อกำหนดและเงื่อนไขแห่งสัญญานี้รวมทั้งเอกสารแนบท้ายสัญญาผนวก .......ทั้งนี้ ที่ปรึกษาจะต้องปฏิบัติงานให้เป็นไปตามหลักวิชาการและมาตรฐานวิชาชีพทางด้าน "+result.ConsultExpertise +" และบทบัญญัติแห่งกฎหมาย ที่เกี่ยวข้อง", null, "32"));
                    string strProjectStartDate = CommonDAO.ToThaiDateStringCovert(result.ProjectStartDate ?? DateTime.Now);
                    string strProjectEndDate = CommonDAO.ToThaiDateStringCovert(result.ProjectEndDate ?? DateTime.Now);
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("1.2 ที่ปรึกษาจะต้องเริ่มลงมือทำงานภายใน"+ strProjectStartDate + " และจะต้องดำเนินการตามสัญญานี้ให้แล้วเสร็จภายใน"+ strProjectEndDate + "", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 2เอกสารอันเป็นส่วนหนึ่งของสัญญา", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เอกสารแนบท้ายสัญญาดังต่อไปนี้ ให้ถือเป็นส่วนหนึ่งของสัญญานี้", null, "32"));

                    //เอกสารแนบท้ายสัญญา
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("2.1 ผนวก 1 …(ขอบข่ายของงานและกำหนดระยะเวลาการทำงาน)… จำนวน....(.…)หน้า", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("2.2 ผนวก 2 …(กำหนดระยะเวลาการทำงานของที่ปรึกษา)… จำนวน....(….)หน้า", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("2.3 ผนวก 3 ...(ค่าจ้างและวิธีการจ่ายค่าจ้าง)… จำนวน….(….)หน้า", null, "32"));
                    //เอกสารแนบท้ายสัญญา

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ความใดในเอกสารแนบท้ายสัญญาที่ขัดหรือแย้งกับข้อความในสัญญานี้ ให้ใช้ข้อความ ในสัญญานี้บังคับ และในกรณีที่เอกสารแนบท้ายสัญญาขัดแย้งกันเอง ที่ปรึกษาจะต้องปฏิบัติตามคำวินิจฉัยของ ผู้ว่าจ้าง คำวินิจฉัยของผู้ว่าจ้างให้ถือเป็นที่สุด และที่ปรึกษาไม่มีสิทธิเรียกร้องค่าจ้าง ค่าเสียหาย หรือค่าใช้จ่ายใดๆ เพิ่มเติมจากผู้ว่าจ้างทั้งสิ้น", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ถ้าสิ่งใดหรือการอันหนึ่งอันใดที่มิได้ระบุไว้ในรายการละเอียดแนบท้ายสัญญานี้ แต่เป็นการอันจำเป็นต้องทำเพื่อให้งานแล้วเสร็จบริบูรณ์ถูกต้องหรือบรรลุผลตามวัตถุประสงค์แห่งสัญญานี้ ที่ปรึกษาต้องจัดทำการนั้นๆ ให้โดยไม่คิดเอาค่าเสียหาย ค่าใช้จ่ายหรือค่าตอบแทนเพิ่มเติมใดๆ ทั้งสิ้น", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 3ค่าจ้างและการจ่ายเงิน", null, "32", true));
                    string strContractTotalAmount =CommonDAO.NumberToThaiText(result.ContractTotalAmount ?? 0);
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ว่าจ้างและที่ปรึกษาได้ตกลงราคาค่าจ้างตามสัญญานี้ เป็นจำนวนเงินทั้งสิ้น "+ result.ContractTotalAmount + "บาท ("+ strContractTotalAmount + ")ซึ่งได้รวมภาษีมูลค่าเพิ่ม เป็นเงินจำนวน…………...…….……..บาท (………..………………….…)ตลอดจนภาษีอากรอื่นๆ และค่าใช้จ่ายทั้งปวงด้วยแล้ว", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ค่าจ้างจะแบ่งออกเป็น "+result.ContractInstallment+"งวด ซึ่งแต่ละงวดจะจ่ายให้เมื่อที่ปรึกษาได้ปฏิบัติงานตามที่กำหนดในเอกสารแนบท้ายสัญญาผนวก .......และคณะกรรมการตรวจรับพัสดุได้พิจารณาแล้วเห็นว่าครบถ้วนถูกต้องและตรวจรับเรียบร้อยแล้ว", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ว่าจ้างอาจจะยึดหน่วงเงินค่าจ้างงวดใดๆ ไว้ก็ได้ หากที่ปรึกษาปฏิบัติงานไม่เป็นไปตามสัญญา และจะจ่ายให้ต่อเมื่อที่ปรึกษาได้ทำการแก้ไขข้อบกพร่องนั้นแล้ว", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(7)การจ่ายเงินตามเงื่อนไขแห่งสัญญานี้ ผู้ว่าจ้างจะโอนเงินเข้าบัญชีเงินฝากธนาคารของ ที่ปรึกษา ชื่อธนาคาร "+result.ContractBankName+" สาขา"+result.ContractBankBranch+ "ชื่อบัญชี "+result.ContractBankAccountName+" เลขที่บัญชี "+result.ContractBankAccountNumber+" ทั้งนี้ ที่ปรึกษาตกลงเป็น" +
                        "ผู้รับภาระเงินค่าธรรมเนียมหรือค่าบริการอื่นใดเกี่ยวกับการโอน รวมทั้งค่าใช้จ่ายอื่นใด (ถ้ามี)ที่ธนาคารเรียกเก็บและยินยอมให้มีการหักเงินดังกล่าวจากจำนวนเงินโอนในงวดนั้นๆ (ความในวรรคนี้ใช้สำหรับกรณีที่หน่วยงานของรัฐจะจ่ายเงินตรงให้แก่ที่ปรึกษา (ระบบ Direct Payment)โดยการโอนเงินเข้าบัญชีเงินฝากธนาคารของที่ปรึกษา ตามแนวทางที่กระทรวงการคลังหรือหน่วยงานของรัฐเจ้าของงบประมาณเป็นผู้กำหนด แล้วแต่กรณี)", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ค่าใช้จ่ายส่วนที่เบิกคืนได้(ถ้ามี)ผู้ว่าจ้างจะจ่ายคืนให้แก่ที่ปรึกษาสำหรับค่าใช้จ่ายซึ่งที่ปรึกษาได้ใช้จ่ายไปตามความเป็นจริงตามเงื่อนไขที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก .....", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(8)ข้อ 4 เงินค่าจ้างล่วงหน้า", null, "32", true));
                    string strPrepaidAmount = CommonDAO.NumberToThaiText(result.PrepaidAmount ?? 0);
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ว่าจ้างตกลงจ่ายเงินค่าจ้างล่วงหน้าให้แก่ที่ปรึกษา เป็นจำนวนเงิน "+result.PrepaidAmount+" บาท("+ strPrepaidAmount + ")ซึ่งเท่ากับร้อยละ "+result.PrepaidPercents+" ของค่าจ้างตามสัญญา", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เงินค่าจ้างล่วงหน้าดังกล่าวจะจ่ายให้ภายหลังจากที่ที่ปรึกษาได้วางหลักประกันการรับเงินค่าจ้างล่วงหน้าเป็น "+result.GuaranteeType+" (หนังสือค้ำประกันหรือหนังสือค้ำประกันอิเล็กทรอนิกส์ของธนาคารภายในประเทศ …………………………....เต็มตามจำนวนเงินค่าจ้างล่วงหน้านั้นให้แก่ผู้ว่าจ้าง ที่ปรึกษาจะต้องออกใบเสร็จรับเงินค่าจ้างล่วงหน้าตามแบบที่ผู้ว่าจ้างกำหนดให้และที่ปรึกษาตกลงที่จะกระทำตามเงื่อนไขอันเกี่ยวกับการใช้จ่ายและการใช้คืนเงินค่าจ้างล่วงหน้านั้น ดังต่อไปนี้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("4.1.ที่ปรึกษาจะใช้เงินค่าจ้างล่วงหน้านั้นเพื่อเป็นค่าใช้จ่ายในการปฏิบัติงานตามสัญญาเท่านั้นหากที่ปรึกษาใช้จ่ายเงินค่าจ้างล่วงหน้าหรือส่วนใดส่วนหนึ่งของเงินค่าจ้างล่วงหน้านั้นในทางอื่นผู้ว่าจ้างอาจจะเรียกเงินค่าจ้างล่วงหน้านั้นคืนจากที่ปรึกษาหรือบังคับเอาจากหลักประกันการรับเงินค่าจ้างล่วงหน้าได้ทันที", JustificationValues.Left, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("4.2.เมื่อผู้ว่าจ้างเรียกร้อง ที่ปรึกษาต้องแสดงหลักฐานการใช้จ่ายเงินค่าจ้างล่วงหน้า เพื่อพิสูจน์ว่าได้เป็นไปตามข้อ ๔.๑ ภายในกำหนด 15 (สิบห้า)วัน นับถัดจากวันได้รับแจ้งจากผู้ว่าจ้าง หากที่ปรึกษาไม่อาจแสดงหลักฐานดังกล่าว ภายในกำหนด 15 (สิบห้า)วัน ผู้ว่าจ้างอาจเรียกเงินค่าจ้างล่วงหน้าคืนจากที่ปรึกษาหรือบังคับเอาจากหลักประกันการรับเงินค่าจ้างล่วงหน้าได้ทันที", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("4.3.ในการจ่ายเงินค่าจ้างให้แก่ที่ปรึกษาตามข้อ 3 ผู้ว่าจ้างจะหักชดใช้คืนเงินค่าจ้างล่วงหน้าในแต่ละงวดไว้จำนวนร้อยละ "+result.PrepaidDeductPercent+" ของจำนวนเงินค่าจ้างในแต่ละงวดจนกว่าจำนวนเงินที่หักไว้จะครบตามจำนวนเงินที่หักค่าจ้างล่วงหน้าที่ที่ปรึกษาได้รับไปแล้ว ยกเว้นค่าจ้างงวดสุดท้ายจะหักไว้เป็นจำนวนเท่ากับจำนวนเงินค่าจ้างล่วงหน้าที่เหลือทั้งหมด", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("4.4.เงินจำนวนใดๆ ก็ตามที่ที่ปรึกษาจะต้องจ่ายให้แก่ผู้ว่าจ้างเพื่อชำระหนี้หรือเพื่อชดใช้ความรับผิดต่างๆ ตามสัญญา ผู้ว่าจ้างจะหักเอาจากเงินค่าจ้างงวดที่จะจ่ายให้แก่ที่ปรึกษาก่อนที่จะหักชดใช้คืนเงินค่าจ้างล่วงหน้า", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("4.5.ในกรณีที่มีการบอกเลิกสัญญา หากเงินค่าจ้างล่วงหน้าที่เหลือเกินกว่าจำนวนเงินที่ ที่ปรึกษาจะได้รับหลังจากหักชดใช้ในกรณีอื่นแล้ว ที่ปรึกษาจะต้องจ่ายคืนเงินจำนวนที่เหลือนั้นให้แก่ผู้ว่าจ้าง ภายใน 7 (เจ็ด)วัน นับถัดจากวันได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้าง", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("4.6.ผู้ว่าจ้างจะคืนหลักประกันเงินค่าจ้างล่วงหน้าให้แก่ที่ปรึกษาต่อเมื่อผู้ว่าจ้างได้หักเงินค่าจ้างไว้จนครบจำนวนเงินค่าจ้างล่วงหน้าตามข้อ ๔.๓ แล้ว", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 5ความรับผิดชอบของที่ปรึกษา", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("5.1.ที่ปรึกษาจะต้องส่งมอบผลงานตามรูปแบบและวิธีการ "+result.SendWorkMethod+" จำนวน "+result.WorkAmount+" ชุด ให้แก่ผู้ว่าจ้างตามที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก1", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("5.2.ในกรณีที่ผลงานของที่ปรึกษาบกพร่องหรือไม่เป็นไปตามข้อกำหนดและเงื่อนไข ตามสัญญาหรือมิได้ดำเนินการให้ถูกต้องตามหลักวิชาการ หรือวิชาชีพ "+result.RelateExpertise+" และ/หรือบทบัญญัติแห่งกฎหมายที่เกี่ยวข้อง ที่ปรึกษาต้องรีบทำการแก้ไขให้เป็นที่เรียบร้อย โดยไม่คิดค่าจ้าง ค่าเสียหาย หรือค่าใช้จ่ายใดๆ จากผู้ว่าจ้างอีก ถ้าที่ปรึกษาหลีกเลี่ยงหรือไม่รีบจัดการแก้ไขให้เป็นที่เรียบร้อยในกำหนดเวลาที่ผู้ว่าจ้างแจ้งเป็นหนังสือ ผู้ว่าจ้างมีสิทธิจ้างที่ปรึกษารายอื่นทำการแทน โดยที่ปรึกษาจะต้องรับผิดชอบจ่ายเงินค่าจ้างในการนี้แทนผู้ว่าจ้างโดยสิ้นเชิง", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ถ้ามีความเสียหายเกิดขึ้นจากงานตามสัญญานี้ไม่ว่าจะเนื่องมาจากการที่ที่ปรึกษา ได้ปฏิบัติงานไม่ถูกต้องตามหลักวิชาการ หรือวิชาชีพ"+result.RelateExpertise+"และ/หรือบทบัญญัติแห่งกฎหมายที่เกี่ยวข้อง หรือเหตุใด ที่ปรึกษาจะต้องทำการแก้ไขความเสียหายดังกล่าว ภายในเวลาที่ผู้ว่าจ้างกำหนดให้ ถ้าที่ปรึกษา ไม่สามารถแก้ไขได้ ที่ปรึกษาจะต้องชดใช้ค่าเสียหายที่เกิดขึ้นแก่ผู้ว่าจ้างโดยสิ้นเชิง ซึ่งรวมทั้งความเสียหายที่เกิดขึ้นโดยตรง และโดยส่วนที่เกี่ยวเนื่องกับความเสียหายที่เกิดขึ้นจากงานตามสัญญานี้ด้วย", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("การที่ผู้ว่าจ้างได้ให้การรับรองหรือความเห็นชอบหรือความยินยอมใดๆ ในการปฏิบัติงานหรือผลงานของที่ปรึกษาหรือการชำระเงินค่าจ้างตามสัญญาแก่ที่ปรึกษา ไม่เป็นการปลดเปลื้องพันธะและ ความรับผิดชอบใดๆ ของที่ปรึกษาตามสัญญานี้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("5.3.บุคลากรหลักของที่ปรึกษา ต้องมีระยะเวลาปฏิบัติงานตามสัญญานี้ไม่ซ้ำซ้อนกับงานในโครงการอื่นๆ ของที่ปรึกษาที่ดำเนินการในช่วงเวลาเดียวกัน หากผู้ว่าจ้างพบว่าบุคลากรหลักไม่ว่าคนหนึ่งคนใดหรือหลายคนปฏิบัติงานซ้ำซ้อนกับงานในโครงการอื่นๆ ไม่ว่าจะพบในระหว่างปฏิบัติงานตามสัญญาหรือในภายหลัง ผู้ว่าจ้างมีสิทธิบอกเลิกสัญญา และ/หรือเรียกค่าเสียหายจากที่ปรึกษาหรือปรับลดค่าจ้างได้", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 6การระงับการทำงานชั่วคราวและการบอกเลิกสัญญา", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("6.1.ผู้ว่าจ้างมีสิทธิบอกเลิกสัญญาในกรณีดังต่อไปนี้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(ก)หากผู้ว่าจ้างเห็นว่าที่ปรึกษามิได้ปฏิบัติงานด้วยความชำนาญหรือด้วยความเอา" +
                        "ใจใส่ในวิชาชีพของที่ปรึกษาเท่าที่พึงคาดหมายได้จากที่ปรึกษาในระดับเดียวกัน หรือมิได้ปฏิบัติตามสัญญาข้อใดข้อหนึ่ง ในกรณีเช่นนี้ผู้ว่าจ้างจะบอกกล่าวให้ที่ปรึกษาทราบถึงเหตุผลที่จะ" +
                        "บอกเลิกสัญญา ถ้าที่ปรึกษามิได้ดำเนินการแก้ไขให้ผู้ว่าจ้างพอใจภายในระยะเวลา "+result.FixDaysAfterNoti+" วัน นับถัดจากวันที่ได้รับคำบอกกล่าว " +
                        "ผู้ว่าจ้างมีสิทธิบอกเลิกสัญญาโดยการส่งคำบอกกล่าวแก่ที่ปรึกษา เมื่อที่ปรึกษาได้รับหนังสือบอกกล่าวนั้นแล้ว ที่ปรึกษาต้องหยุดปฏิบัติงานทันที และดำเนินการทุกวิถีทางเพื่อลดค่าใช้จ่ายใดๆ " +
                        "ที่อาจมีในระหว่างการหยุดปฏิบัติงานนั้นให้น้อยที่สุด", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(ข)หากที่ปรึกษามิได้ลงมือทำงานภายในกำหนดเวลา " +
                        "หรือไม่สามารถทำงานให้แล้วเสร็จตามกำหนดเวลา หรือมีเหตุให้ผู้ว่าจ้างเชื่อได้ว่าที่ปรึกษาไม่สามารถทำงานให้แล้วเสร็จภายในกำหนดเวลา" +
                        " หรือล่วงเลยกำหนดเวลาแล้วเสร็จไปแล้ว หรือตกเป็นผู้ล้มละลาย ผู้ว่าจ้างมีสิทธิบอกเลิกสัญญาได้ทันที", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("การบอกเลิกสัญญาตามข้อ ๖.๑ ผู้ว่าจ้างมีสิทธิริบหรือบังคับจากหลักประกันเงินค่าจ้างล่วงหน้า หลักประกันการปฏิบัติตามสัญญา" +
                        "และเงินประกันผลงานทั้งหมดหรือแต่บางส่วน และมีสิทธิเรียกค่าเสียหายอื่น (ถ้ามี)จากที่ปรึกษาด้วย", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("6.2.ผู้ว่าจ้างอาจมีหนังสือบอกกล่าวให้ที่ปรึกษาทราบล่วงหน้าเมื่อใดก็ได้ว่าผู้ว่าจ้างมีเจตนาที่จะระงับการทำงานของที่ปรึกษา" +
                        "ไว้ชั่วคราวไม่ว่าทั้งหมดหรือแต่บางส่วน หรือจะบอกเลิกสัญญา ในกรณีที่ผู้ว่าจ้าง จะบอกเลิกสัญญา การบอกเลิกสัญญาดังกล่าวจะมีผลในเวลาไม่น้อยกว่า "+result.NotiDaysAfterTerminate+" วัน นับถัดจากวันที่ที่ปรึกษาได้รับหนังสือบอกกล่าวนั้น หรืออาจเร็วกว่าหรือช้ากว่ากำหนดเวลานั้นก็ได้แล้วแต่คู่สัญญาจะทำความตกลงกัน " +
                        "เมื่อที่ปรึกษาได้รับหนังสือบอกกล่าวนั้นแล้ว ที่ปรึกษาต้องหยุดปฏิบัติงานทันที ทั้งนี้ ที่ปรึกษาไม่มีสิทธิได้รับค่าจ้างในระหว่างระงับการทำงานไว้ชั่วคราว และที่ปรึกษาจะต้องดำเนินการทุกวิถีทางเพื่อลดค่าใช้จ่ายใดๆ ที่อาจมีในระหว่างการหยุดปฏิบัติงานนั้นให้น้อยที่สุด", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("กรณีที่มีการระงับการทำงานชั่วคราวตามข้อ ๖.๒ ผู้ว่าจ้างจะจ่ายเงินเป็นค่าใช้จ่ายเท่าที่จำเป็นให้แก่ที่ปรึกษาตามที่ผู้ว่าจ้างเห็นสมควร", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("กรณีที่มีการเลิกสัญญาตามข้อ ๖.๒ ผู้ว่าจ้างจะชำระค่าจ้างตามส่วนที่เป็นธรรมและเหมาะสมที่กำหนดในเอกสารแนบท้ายสัญญาผนวก ......ให้แก่ที่ปรึกษา โดยคำนวณตั้งแต่วันเริ่มปฏิบัติงานจนถึงวันบอกเลิกสัญญา นอกจากนี้ผู้ว่าจ้างจะคืนหลักประกันการปฏิบัติตามสัญญาหรือเงินประกันผลงาน " +
                        "รวมทั้งเงินชดเชยค่าเดินทางและเงินค่าใช้จ่ายที่ได้ทดรองจ่ายไปตามสมควรและตามความเป็นจริง ซึ่งผู้ว่าจ้างยังมิได้ชำระให้แก่ที่ปรึกษาด้วย อย่างไรก็ตาม เงินชดเชยและเงินที่ได้ชำระไปแล้วทั้งหมดจะต้องไม่เกินราคาค่าจ้างตามข้อ ๓", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 7สิทธิและหน้าที่ของที่ปรึกษา", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("7.1.ที่ปรึกษาจะต้องใช้ความชำนาญ ความระมัดระวัง และความขยันหมั่นเพียร ในการปฏิบัติงานตามสัญญาอย่างมีประสิทธิภาพ และจะต้องปฏิบัติหน้าที่ตามความรับผิดชอบให้สำเร็จลุล่วง เป็นไปตามมาตรฐานของวิชาชีพที่ยอมรับนับถือกันโดยทั่วไป", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("7.2.ค่าจ้างซึ่งผู้ว่าจ้างจะชำระแก่ที่ปรึกษาตามข้อ ๓ เป็นค่าตอบแทนเพียงอย่างเดียวเท่านั้นซึ่งที่ปรึกษาจะได้รับเกี่ยวกับการปฏิบัติงานตามสัญญานี้ ที่ปรึกษาจะต้องไม่รับค่านายหน้าทางการค้า ส่วนลด เบี้ยเลี้ยง เงินช่วยเหลือหรือผลประโยชน์ใดๆ ไม่ว่าโดยตรงหรือโดยอ้อม หรือสิ่งตอบแทนใดๆในส่วนที่เกี่ยวข้องกับสัญญานี้ หรือที่เกี่ยวกับการปฏิบัติหน้าที่ตามสัญญานี้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("7.3.ที่ปรึกษาจะต้องไม่มีผลประโยชน์ใดๆ ไม่ว่าโดยตรงหรือโดยอ้อมในเงินค่าสิทธิ เงินบำเหน็จ หรือค่านายหน้าใดๆ ที่เกี่ยวกับการนำวัสดุสิ่งของหรือกรรมวิธีใดๆ ที่มีทะเบียนสิทธิบัตรหรือได้รับการคุ้มครองทางทรัพย์สินทางปัญญาหรือตามกฎหมายอื่นใดมาใช้เพื่อวัตถุประสงค์ของสัญญานี้ เว้นแต่คู่สัญญา ทั้งสองฝ่ายจะได้ตกลงกันเป็นหนังสือว่าที่ปรึกษาอาจจะได้ผลประโยชน์หรือเงินเช่นว่านั้นได้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("7.4.บรรดางานและเอกสารที่ที่ปรึกษาได้จัดทำขึ้นเกี่ยวกับสัญญานี้ให้ถือเป็นความลับและให้ตกเป็นกรรมสิทธิ์ของผู้ว่าจ้าง ที่ปรึกษาจะต้องส่งมอบบรรดางานและเอกสารดังกล่าวให้แก่ผู้ว่าจ้างเมื่อสิ้นสุดสัญญานี้ ที่ปรึกษาอาจเก็บสำเนาเอกสารไว้กับตนได้แต่ต้องไม่นำข้อความในเอกสารนั้นไปใช้ในกิจการอื่นที่ ไม่เกี่ยวกับงานโดยไม่ได้รับความยินยอมล่วงหน้าเป็นหนังสือจากผู้ว่าจ้างก่อน", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("7.5.ผู้ว่าจ้างเป็นเจ้าของลิขสิทธิ์หรือสิทธิในทรัพย์สินทางปัญญา รวมถึงสิทธิใดๆ ในผลงานที่ที่ปรึกษาได้ปฏิบัติงานตามสัญญานี้แต่เพียงฝ่ายเดียว และที่ปรึกษาจะนำผลงาน และ/หรือรายละเอียดของงานตามสัญญานี้ ไม่ว่าทั้งหมดหรือบางส่วนไปใช้ หรือเผยแพร่ในกิจการอื่น นอกเหนือจากที่ได้ระบุไว้ในสัญญานี้ไม่ได้ เว้นแต่ได้รับอนุญาตเป็นหนังสือจากผู้ว่าจ้างก่อน", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("7.6.บรรดาเครื่องมือ เครื่องใช้ และวัสดุอุปกรณ์ทั้งหลาย ซึ่งผู้ว่าจ้างได้จัดให้ที่ปรึกษาใช้หรือซึ่งที่ปรึกษาซื้อมาด้วยทุนทรัพย์ของผู้ว่าจ้าง หรือซึ่งผู้ว่าจ้างเป็นผู้จ่ายชดใช้คืนให้ ถือว่าเป็นกรรมสิทธิ์ของผู้ว่าจ้างและต้องทำข้อความและเครื่องหมายที่แสดงว่าเป็นของผู้ว่าจ้างไว้ที่ทรัพย์สินดังกล่าวด้วย ทั้งนี้ ที่ปรึกษาต้องใช้เครื่องมือเครื่องใช้และวัสดุอุปกรณ์ดังกล่าวอย่างเหมาะสมตามระเบียบของผู้ว่าจ้างหรือของทางราชการ ที่เกี่ยวข้องเพื่อกิจการที่เกี่ยวกับการจ้างที่ปรึกษาเท่านั้น", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เมื่อที่ปรึกษาทำงานเสร็จหรือมีการเลิกสัญญา ที่ปรึกษาจะต้องทำบัญชีแสดงรายการเครื่องมือเครื่องใช้และวัสดุอุปกรณ์ทั้งหลายข้างต้นที่ยังคงเหลืออยู่ และจัดการโยกย้ายไปเก็บรักษาตามคำสั่งผู้ว่าจ้างโดยพลัน ที่ปรึกษาต้องดูแลเครื่องมือเครื่องใช้และวัสดุอุปกรณ์ดังกล่าวอย่างเหมาะสมตลอดเวลาที่ครอบครอง และต้องคืนเครื่องมือเครื่องใช้และวัสดุอุปกรณ์ดังกล่าวให้ครบถ้วนในสภาพดีตามความเหมาะสม แต่ไม่ต้องรับผิดชอบสำหรับความเสื่อมสภาพจากการใช้งานตามปกติ", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("7.7.ที่ปรึกษาจะจัดให้มีบุคลากรที่มีความรู้และความชำนาญงานมาปฏิบัติงานให้เหมาะสมกับสภาพการปฏิบัติงานตามสัญญานี้และให้สอดคล้องกับขอบเขตของงานของที่ปรึกษาตามที่ปรากฏ ในเอกสารแนบท้ายสัญญาผนวก …..การเปลี่ยนแปลงบุคลากรดังกล่าวจะต้องได้รับความยินยอมเป็นหนังสือจากผู้ว่าจ้างก่อน", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("7.8.ในกรณีที่ผู้ว่าจ้างพิจารณาเห็นว่า การดำเนินงานของบุคลากรที่ที่ปรึกษาจัดหามาจะก่อให้เกิดความเสียหายแก่งานตามสัญญานี้ ไม่ว่าในกรณีใดก็ตามผู้ว่าจ้างมีสิทธิที่จะให้ที่ปรึกษาเปลี่ยนบุคลากรบางคนหรือทั้งหมดนั้นได้ และที่ปรึกษาต้องดำเนินการตามความประสงค์ของผู้ว่าจ้างโดยเร็ว", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("การเปลี่ยนบุคลากรตามความในวรรคก่อน ที่ปรึกษาจะต้องเสนอรายชื่อบุคลากรที่จะปฏิบัติงานแทนนั้น ต่อผู้ว่าจ้างเพื่อพิจารณาให้ความเห็นชอบก่อน", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 8ความรับผิดชอบของที่ปรึกษาต่อบุคคลภายนอก", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("8.1.ที่ปรึกษาจะต้องชดใช้ค่าเสียหายให้แก่ผู้ว่าจ้าง และป้องกันมิให้ผู้ว่าจ้างต้องรับผิดชอบในบรรดาสิทธิเรียกร้อง ค่าเสียหาย ค่าใช้จ่าย หรือราคา รวมตลอดถึงการเรียกร้องโดยบุคคลภายนอกอันเกิดจากความผิดพลาดหรือการละเว้นไม่กระทำการของที่ปรึกษา หรือของลูกจ้างของที่ปรึกษา", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("8.2.ที่ปรึกษาจะต้องรับผิดชอบต่อการละเมิดบทบัญญัติแห่งกฎหมาย หรือการละเมิดลิขสิทธิ์ หรือสิทธิในทรัพย์สินทางปัญญาอื่น รวมถึงสิทธิใดๆ ต่อบุคคลภายนอกเนื่องจากการปฏิบัติงานตามสัญญานี้ โดยสิ้นเชิง", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(9)8.3.ที่ปรึกษาจะต้องจัดการประกันภัยกับบริษัทประกันภัยที่ผู้ว่าจ้างเห็นชอบเพื่อความรับผิดต่อบุคคลภายนอก และเพื่อความสูญหายหรือเสียหายในทรัพย์สินซึ่งผู้ว่าจ้างเป็นผู้จัดหาให้หรือสั่งซื้อโดยทุนทรัพย์ของผู้ว่าจ้าง เพื่อให้ที่ปรึกษาไว้ใช้ในการปฏิบัติงานตามสัญญานี้ โดยที่ปรึกษาเป็นผู้ออกค่าใช้จ่ายในการประกันภัยเอง ทั้งนี้ เว้นแต่จะมีการตกลงกันไว้เป็นอย่างอื่นในสัญญานี้", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 9พันธะหน้าที่ของผู้ว่าจ้าง", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ว่าจ้างจะมอบข้อมูลและสถิติต่างๆ ที่เกี่ยวข้องซึ่งผู้ว่าจ้างมีอยู่ให้แก่ที่ปรึกษาโดยไม่คิดมูลค่าและภายในเวลาอันควร", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ที่ปรึกษาร้องขอความช่วยเหลือ ผู้ว่าจ้างจะพิจารณาให้ความช่วยเหลืออำนวยความสะดวกตามสมควร ทั้งนี้ เพื่อให้การปฏิบัติงานของที่ปรึกษาตามสัญญานี้ลุล่วงไปได้ด้วยดี", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 10ค่าปรับ", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ที่ปรึกษาไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่งด้วยเหตุใดๆ ก็ตาม จนเป็นเหตุให้เกิดค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายแก่ผู้ว่าจ้าง ที่ปรึกษาต้องชดใช้ค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายดังกล่าวให้แก่ผู้ว่าจ้างโดยสิ้นเชิงภายในกำหนด..........................(............................)วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้าง หากที่ปรึกษาไม่ชดใช้ให้ถูกต้องครบถ้วนภายในระยะเวลาดังกล่าว ให้ผู้ว่าจ้างมีสิทธิที่จะหักเอาจากจำนวนเงินค่าจ้างของที่ปรึกษาที่ต้องชำระ หรือบังคับจากหลักประกันการปฏิบัติตามสัญญา หรือเงินประกันผลงานได้ทันที", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หากค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายที่บังคับจากเงินค่าจ้างที่ต้องชำระ หลักประกัน การปฏิบัติตามสัญญา และเงินประกันผลงานแล้วยังไม่เพียงพอ ที่ปรึกษายินยอมชำระส่วนที่เหลือที่ยังขาดอยู่จนครบถ้วนตามจำนวนค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายนั้น ภายในกำหนด..................(....................)วันนับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้าง", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หากมีเงินค่าจ้างตามสัญญาที่หักไว้จ่ายเป็นค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายแล้วยังเหลืออยู่อีกเท่าใด ผู้ว่าจ้างจะคืนให้แก่ที่ปรึกษาทั้งหมด", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(11)ข้อ 12 (ก)เงินประกันผลงาน", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในการจ่ายเงินให้แก่ที่ปรึกษาแต่ละงวด ผู้ว่าจ้างจะหักเงินจำนวนร้อยละ…(12)….(................)ของเงินที่ต้องจ่ายในงวดนั้นเพื่อเป็นประกันผลงาน หรือที่ปรึกษาอาจนำหนังสือค้ำประกันของธนาคารหรือหนังสือค้ำประกันอิเล็กทรอนิกส์ของธนาคารภายในประเทศซึ่งมีอายุการค้ำประกันตลอดอายุสัญญามามอบให้ผู้ว่าจ้าง ทั้งนี้เพื่อเป็นหลักประกันแทนก็ได้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ว่าจ้างจะคืนเงินประกันผลงาน และ/หรือหนังสือค้ำประกันของธนาคารดังกล่าวตามวรรคหนึ่งโดยไม่มีดอกเบี้ยให้แก่ที่ปรึกษาพร้อมกับการจ่ายเงินค่าจ้างงวดสุดท้าย", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๒ (ข)หลักประกันการปฏิบัติตามสัญญา", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในขณะทำสัญญานี้ที่ปรึกษาได้นำหลักประกันเป็น……........(13)….….....….เป็นจำนวนเงิน…..……….……บาท (………………………)ซึ่งเท่ากับร้อยละ.…(14)....(……………….......)ของราคาค่าจ้างตามสัญญา มามอบไว้แก่ผู้ว่าจ้างเพื่อเป็นหลักประกันการปฏิบัติตามสัญญานี้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(15)กรณีที่ปรึกษาใช้หนังสือค้ำประกันมาเป็นหลักประกันการปฏิบัติตามสัญญา หนังสือค้ำประกันดังกล่าวจะต้องออกโดยธนาคารที่ประกอบกิจการในประเทศไทย หรือโดยบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบ โดยอนุโลมให้ใช้ตามตัวอย่างหนังสือค้ำประกันของธนาคารที่คณะกรรมการนโยบายกำหนด หรืออาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้ และจะต้องมีอายุการค้ำประกันตลอดไปจนกว่าที่ปรึกษาพ้นข้อผูกพันตามสัญญานี้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หลักประกันจะต้องมีอายุครอบคลุมความรับผิดทั้งปวงของที่ปรึกษาตลอดอายุสัญญา ถ้าหลักประกันที่ที่ปรึกษานำมามอบให้ดังกล่าวลดลงหรือเสื่อมค่าลง หรือมีอายุไม่ครอบคลุมถึงความรับผิดของ ที่ปรึกษาตลอดอายุสัญญา ไม่ว่าด้วยเหตุใดๆ ก็ตาม รวมถึงกรณีที่ปรึกษาส่งมอบงานล่าช้าเป็นเหตุให้ระยะเวลา แล้วเสร็จตามสัญญาเปลี่ยนแปลงไป ที่ปรึกษาต้องหาหลักประกันใหม่หรือหลักประกันเพิ่มเติมให้มีจำนวนครบถ้วน ตามวรรคหนึ่งมามอบให้แก่ผู้ว่าจ้างภายใน.............(………..….)วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้าง", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หลักประกันที่ผู้รับจ้างนำมามอบไว้ตามข้อนี้ ผู้ว่าจ้างจะคืนให้แก่ที่ปรึกษาโดยไม่มีดอกเบี้ย เมื่อที่ปรึกษาพ้นจากข้อผูกพันและความรับผิดทั้งปวงตามสัญญานี้แล้ว", null, "32"));


                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๓ การจ้างช่วง", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ที่ปรึกษาจะต้องไม่โอนสิทธิประโยชน์ใดๆ ตามสัญญานี้ให้แก่ผู้อื่นโดยไม่ได้รับความยินยอมเป็นหนังสือจากผู้ว่าจ้างก่อน เว้นแต่การโอนสิทธิที่จะรับเงินค่าจ้างตามสัญญานี้", null, "32"));


                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๕การงดหรือลดค่าปรับ หรือขยายเวลาปฏิบัติงานตามสัญญา", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่มีเหตุเกิดจากความผิดหรือความบกพร่องของผู้ว่าจ้างหรือเหตุสุดวิสัย หรือเกิดจากพฤติการณ์อันหนึ่งอันใดที่ที่ปรึกษาไม่ต้องรับผิดตามกฎหมาย หรือเหตุอื่นตามที่กำหนดในกฎกระทรวง ซึ่งออกตามความในกฎหมายว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ ทำให้ที่ปรึกษาไม่สามารถทำงานให้แล้วเสร็จตามเงื่อนไขและกำหนดเวลาแห่งสัญญานี้ได้ ที่ปรึกษาจะต้องแจ้งเหตุหรือพฤติการณ์ดังกล่าว พร้อมหลักฐานเป็นหนังสือให้ผู้ว่าจ้างทราบ เพื่อของดหรือลดค่าปรับ หรือขยายเวลาทำงานออกไปภายใน ๑๕ (สิบห้า)วัน นับถัดจากวันที่เหตุนั้นสิ้นสุดลง หรือตามที่กำหนดในกฎกระทรวงดังกล่าว แล้วแต่กรณี", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ถ้าที่ปรึกษาไม่ปฏิบัติให้เป็นไปตามความในวรรคหนึ่ง ให้ถือว่าที่ปรึกษาได้สละสิทธิเรียกร้อง ในการที่จะของดหรือลดค่าปรับ หรือขยายเวลาทำงานออกไปโดยไม่มีเงื่อนไขใดๆ ทั้งสิ้น เว้นแต่กรณีเหตุเกิดจากความผิดหรือความบกพร่องของฝ่ายผู้ว่าจ้างซึ่งมีหลักฐานชัดแจ้งหรือผู้ว่าจ้างทราบดีอยู่แล้วตั้งแต่ต้น", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("การงดหรือลดค่าปรับ หรือขยายกำหนดเวลาทำงานตามวรรคหนึ่ง อยู่ในดุลพินิจของผู้ว่าจ้างที่จะพิจารณาตามที่เห็นสมควร", null, "32"));


                    body.AppendChild(WordServiceSetting.EmptyParagraph());
                    body.AppendChild(WordServiceSetting.EmptyParagraph());
                    body.AppendChild(WordServiceSetting.EmptyParagraph());

                    body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ).................................................................ผู้ว่าจ้าง", "32"));
                    body.AppendChild(WordServiceSetting.CenteredParagraph("(....................................................)", "32"));

                    body.AppendChild(WordServiceSetting.EmptyParagraph());
                    body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ).................................................................ที่ปรึกษา", "32"));
                    body.AppendChild(WordServiceSetting.CenteredParagraph("(....................................................)", "32"));

                    body.AppendChild(WordServiceSetting.EmptyParagraph());
                    body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ).................................................................พยาน", "32"));
                    body.AppendChild(WordServiceSetting.CenteredParagraph("(....................................................)", "32"));

                    body.AppendChild(WordServiceSetting.EmptyParagraph());
                    body.AppendChild(WordServiceSetting.CenteredParagraph("(ลงชื่อ).................................................................พยาน", "32"));
                    body.AppendChild(WordServiceSetting.CenteredParagraph("(....................................................)", "32"));

                    // next page
                    body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                    body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("วิธีปฏิบัติเกี่ยวกับสัญญาเช่าเครื่องถ่ายเอกสาร", "000000", "36"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(1)ให้ระบุเลขที่สัญญาในปีงบประมาณหนึ่งๆ ตามลำดับ", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(2)ให้ระบุชื่อของหน่วยงานของรัฐที่เป็นนิติบุคคล เช่น กรม ก.หรือรัฐวิสาหกิจ ข.เป็นต้น", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(3)ให้ระบุชื่อและตำแหน่งของหัวหน้าหน่วยงานของรัฐที่เป็นนิติบุคคลนั้น หรือผู้ที่ได้รับมอบอำนาจ เช่น นาย ก.อธิบดีกรม………...… หรือ นาย ข.ผู้ได้รับมอบอำนาจจากอธิบดีกรม………......………..", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(4)ให้ระบุชื่อผู้ให้เช่า", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ก.กรณีนิติบุคคล เช่น ห้างหุ้นส่วนสามัญจดทะเบียน ห้างหุ้นส่วนจำกัด บริษัทจำกัด", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ข.กรณีบุคคลธรรมดา ให้ระบุชื่อและที่อยู่", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(5)เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(6)เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(7)หน่วยงานของรัฐอาจกำหนดเงื่อนไขการจ่ายค่าเช่าให้แตกต่างไปจากแบบสัญญาที่กำหนดได้ตามความเหมาะสมและจำเป็นและไม่ทำให้หน่วยงานของรัฐเสียเปรียบ หากหน่วยงานของรัฐเห็นว่าจะมีปัญหาในทางเสียเปรียบหรือไม่รัดกุมพอ ก็ให้ส่งร่างสัญญานั้นไปให้สำนักงานอัยการสูงสุดพิจารณาให้ความเห็นชอบก่อน", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(8)เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(9)ชื่อสถานที่หน่วยงานของรัฐ", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(10)ให้พิจารณาถึงความจำเป็นและเหมาะสมของการใช้งาน", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(11)อัตราค่าปรับตามสัญญาข้อ 9 ให้กำหนดเป็นรายวันในอัตราระหว่างร้อยละ 0.01 – 0.20 ตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ.2560 ข้อ 162 ส่วนกรณีจะปรับร้อยละเท่าใด ให้อยู่ในดุลยพินิจของหน่วยงานของรัฐผู้เช่าที่จะพิจารณา โดยคำนึงถึงราคาและลักษณะของพัสดุที่เช่า ซึ่งอาจมีผลกระทบต่อการที่ผู้ให้เช่าจะหลีกเลี่ยงไม่ปฏิบัติตามสัญญา แต่ทั้งนี้การที่จะกำหนดค่าปรับเป็นร้อยละเท่าใดจะต้องกำหนดไว้ในเอกสารเชิญชวนด้วย", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(12)“หลักประกัน” หมายถึง หลักประกันที่ผู้ให้เช่านำมามอบไว้แก่หน่วยงานของรัฐ เมื่อลงนามในสัญญา เพื่อเป็นการประกันความเสียหายที่อาจจะเกิดขึ้นจากการปฏิบัติตามสัญญา ดังนี้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๑)เงินสด", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๒)เช็คหรือดราฟท์ ที่ธนาคารเซ็นสั่งจ่าย ซึ่งเป็นเช็คหรือดราฟท์ลงวันที่ที่ใช้เช็คหรือดราฟท์นั้นชำระต่อเจ้าหน้าที่ หรือก่อนวันนั้นไม่เกิน ๓ วันทำการ", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๓)หนังสือคํ้าประกันของธนาคารภายในประเทศตามตัวอย่างที่คณะกรรมการนโยบายกําหนด โดยอาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๔)หนังสือค้ำประกันของบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบ โดยอนุโลมให้ใช้ตามตัวอย่างหนังสือค้ำประกันของธนาคารที่คณะกรรมการนโยบายกำหนด", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๕)พันธบัตรรัฐบาลไทย", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(13)ให้กำหนดจำนวนเงินหลักประกันการปฏิบัติตามสัญญาตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ.2560 ข้อ 168", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(14)เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(15)กำหนดระยะเวลาตามความเหมาะสม เช่น 3 เดือน", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(16)อัตราค่าปรับตามสัญญาข้อ 12 ให้กำหนดเป็นรายวันในอัตราระหว่างร้อยละ 0.01 – 0.20 ตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ.2560 ข้อ 162 ส่วนกรณีจะปรับร้อยละเท่าใด ให้อยู่ในดุลยพินิจของหน่วยงานของรัฐผู้เช่าที่จะพิจารณา โดยคำนึงถึงราคาและลักษณะของพัสดุที่เช่า ซึ่งอาจมีผลกระทบต่อการที่ผู้ให้เช่าจะหลีกเลี่ยงไม่ปฏิบัติตามสัญญา แต่ทั้งนี้การที่จะกำหนดค่าปรับเป็นร้อยละเท่าใดจะต้องกำหนดไว้ในเอกสารเชิญชวนด้วย", null, "32"));


                    body.AppendChild(WordServiceSetting.EmptyParagraph());




                    WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);

                }
                stream.Position = 0;
                return stream.ToArray();
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Error in WordEContract_ConsultantService.OnGetWordContact_ConsultantService: " + ex.Message, ex);
        }
   
    }

    public async Task<string> OnGetWordContact_ConsultantService_ToPDF(string id,string typeContact)
    {
        try
        {
            var result = await _e.GetCTRAsync(id);
            if (result == null)
            {
                throw new Exception("ไม่พบข้อมูลสัญญาจ้างที่ปรึกษา");
            }
            else
            {

                var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");
                string ToThaiDate(DateTime? date) => CommonDAO.ToThaiDateStringCovert(date ?? DateTime.Now);

                // Helper for Thai number text
                string ToThaiNumberText(decimal? number) => CommonDAO.NumberToThaiText(result.ContractTotalAmount ?? 0);
                string strVatAmount(decimal? number) => CommonDAO.NumberToThaiText(result.VatAmount ?? 0);

                var listDocAtt = await _eContractDAO.GetRelatedDocumentsAsync(id, "CTR31760");
                var htmlDocAtt = listDocAtt != null
                    ? string.Join("", listDocAtt.Select(docItem =>
                        $"<p class='tab3 t-16'>{docItem.DocumentTitle} จำนวน {docItem.PageAmount} หน้า</div>"))
                    : "";
                #region  signlist
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
                            signatureHtml = "<div class='t-16 text-center tab1'>(ลงชื่อ....................)</div>";
                        }
                    }
                    else
                    {
                        signatureHtml = "<div class='t-16 text-center tab1'>(ลงชื่อ....................)</div>";
                    }

                    signatoryHtml.AppendLine($@"
    <div class='sign-single-right'>
        {signatureHtml}
        <div class='t-16 text-center tab1'>({signer?.Signatory_Name})</div>
        <div class='t-16 text-center tab1'>{signer?.BU_UNIT}</div>
    </div>");

                    signatoryHtml.Append(companySealHtml);
                }

                #endregion signlist

                // Build HTML body
                var htmlBody = $@"
<div class='contract'>
    <div class='text-center t-22'><b>แบบสัญญา</b></div>
    <div class='text-center t-22'><b>สัญญาจ้างผู้เชี่ยวชาญรายบุคคลหรือจ้างบริษัทที่ปรึกษา</b></div>
</br>
    <div class='text-right t-16'>สัญญาเลขที่…{result.Contract_Number}.</div>
    <p class='tab3 t-16'>
        สัญญาฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.)
        ตำบล/แขวง ทุ่งสองห้อง อำเภอ/เขต หลักสี่
        จังหวัด กรุงเทพ เมื่อ {ToThaiDate(result.ContractSignDate)}
        ระหว่าง {result.Contract_Organization}
        โดย {result.SignatoryName}
        ซึ่งต่อไปในสัญญานี้เรียกว่า “ผู้ว่าจ้าง” 
ฝ่ายหนึ่ง กับ {result.ContractorName}
    </p>
    {(result.ContractorType == "นิติบุคคล" ? $@"
    <p class='tab3 t-16'>
        ซึ่งจดทะเบียนเป็นนิติบุคคล ณ {result.ContractorName} มี
        สำนักงานใหญ่อยู่เลขที่ {result.ContractorAddressNo} ถนน {result.ContractorStreet} ตำบล/แขวง {result.ContractorSubDistrict}
        อำเภอ/เขต {result.ContractorDistrict} จังหวัด {result.ContractorProvince}
        โดย {result.ContractorSignatoryName}
        มีอำนาจลงนามผูกพันนิติบุคคลปรากฏตามหนังสือรับรองของสำนักงานทะเบียนหุ้นส่วนบริษัท ……………
        ลงวันที่ {ToThaiDate(result.ContractSignDate)} (5)(และหนังสือมอบอำนาจลง {ToThaiDate(result.ContractSignDate)}) แนบท้ายสัญญานี้
    </p>
    " : $@"
    <p class='tab3 t-16'>
        (๖)(ในกรณีที่ผู้รับจ้างเป็นบุคคลธรรมดาให้ใช้ข้อความว่า กับ {result.ContractorName}
        อยู่บ้านเลขที่ {result.ContractorAddressNo} ถนน {result.ContractorStreet} ตำบล/แขวง {result.ContractorSubDistrict}
        อำเภอ/เขต {result.ContractorDistrict} จังหวัด {result.ContractorProvince}
        ผู้ถือบัตรประจำตัวประชาชนเลขที่ {result.CitizenId} ดังปรากฏตามสำเนาบัตรประจำตัวประชาชนแนบท้ายสัญญานี้) ซึ่งต่อไปใน
        สัญญานี้เรียกว่า “ที่ปรึกษา” อีกฝ่ายหนึ่ง กับ…{result.ContractorName}
    </p>
    ")}
    <p class='tab3 t-16'>คู่สัญญาได้ตกลงกันมีข้อความดังต่อไปนี้</p>
    <p class='tab3 t-16'><b>ข้อ ๑ ข้อตกลงว่าจ้าง</b></p>
    <p class='tab3 t-16'>
        ๑.๑ ผู้ว่าจ้างตกลงจ้างและที่ปรึกษาตกลงรับจ้างปฏิบัติงานตามโครงการ {result.ProjectName}: {result.ProjectDesc} {result.ProjectReference} ตามข้อกำหนดและเงื่อนไขแห่งสัญญานี้รวมทั้งเอกสารแนบท้ายสัญญาผนวก .......
ทั้งนี้ ที่ปรึกษาจะต้อง
</br>ปฏิบัติงานให้เป็นไปตามหลักวิชาการและมาตรฐานวิชาชีพทางด้าน {result.ConsultExpertise} และบทบัญญัติแห่งกฎหมาย 
</br>ที่เกี่ยวข้อง
    </p>
    <p class='tab3 t-16'>
        ๑.๒ ที่ปรึกษาจะต้องเริ่มลงมือทำงานภายใน {ToThaiDate(result.ProjectStartDate)} และ
</br>จะต้องดำเนินการตามสัญญานี้ให้แล้วเสร็จภายใน {ToThaiDate(result.ProjectEndDate)}
    </p>
    <p class='tab3 t-16'><b>ข้อ ๒ เอกสารอันเป็นส่วนหนึ่งของสัญญา</b></p>
{htmlDocAtt}
    <p class='tab3 t-16'>
        ความใดในเอกสารแนบท้ายสัญญาที่ขัดหรือแย้งกับข้อความในสัญญานี้ ให้ใช้ข้อความ 
</br>ในสัญญานี้บังคับ และในกรณีที่เอกสารแนบท้ายสัญญาขัดแย้งกันเอง ที่ปรึกษาจะต้องปฏิบัติตามคำวินิจฉัย 
</br>ของผู้ว่าจ้าง คำวินิจฉัยของผู้ว่าจ้างให้ถือเป็นที่สุด และที่ปรึกษาไม่มีสิทธิเรียกร้องค่าจ้าง ค่าเสียหาย หรือค่าใช้
</br>จ่ายใดๆ เพิ่มเติมจากผู้ว่าจ้างทั้งสิ้น
    </p>
    <p class='tab3 t-16'>
        ถ้าสิ่งใดหรือการอันหนึ่งอันใดที่มิได้ระบุไว้ในรายการละเอียดแนบท้ายสัญญานี้ แต่เป็น
</br>การอันจำเป็นต้องทำเพื่อให้งานแล้วเสร็จบริบูรณ์ถูกต้องหรือบรรลุผลตามวัตถุประสงค์แห่งสัญญานี้ ที่ปรึกษาต้องจัดทำการนั้นๆ ให้โดยไม่คิดเอาค่าเสียหาย ค่าใช้จ่ายหรือค่าตอบแทนเพิ่มเติมใดๆ ทั้งสิ้น
    </p>
    <p class='tab3 t-16'><b>ข้อ ๓ ค่าจ้างและการจ่ายเงิน</b></p>
    <p class='tab3 t-16'>
        ผู้ว่าจ้างและที่ปรึกษาได้ตกลงราคาค่าจ้างตามสัญญานี้ เป็นจำนวนเงินทั้งสิ้น {result.ContractTotalAmount?.ToString("N2") ?? "0.00"}
 บาท ({ToThaiNumberText(result.ContractTotalAmount)}) ซึ่งได้รวมภาษีมูลค่าเพิ่ม 
เป็นเงินจำนวน{result.VatAmount?.ToString("N2") ?? "0.00"}
บาท ({ToThaiNumberText(result.VatAmount)}) ตลอดจนภาษีอากรอื่นๆ และค่าใช้จ่ายทั้งปวงด้วยแล้ว
    </p>
    <p class='tab3 t-16'>
        ค่าจ้างจะแบ่งออกเป็น {result.ContractInstallment} งวด ซึ่งแต่ละงวดจะจ่ายให้เมื่อที่ปรึกษาได้ปฏิบัติงานตามที่กำหนด
</br>ในเอกสารแนบท้ายสัญญาผนวก .......และคณะกรรมการตรวจรับพัสดุได้พิจารณาแล้วเห็นว่าครบถ้วนถูกต้อง
</br>และตรวจรับเรียบร้อยแล้ว
    </p>
    <p class='tab3 t-16'>
        ผู้ว่าจ้างอาจจะยึดหน่วงเงินค่าจ้างงวดใดๆ ไว้ก็ได้ หากที่ปรึกษาปฏิบัติงานไม่เป็นไปตาม
</br>สัญญา และจะจ่ายให้ต่อเมื่อที่ปรึกษาได้ทำการแก้ไขข้อบกพร่องนั้นแล้ว
    </p>
    <p class='tab3 t-16'>
        (๗)การจ่ายเงินตามเงื่อนไขแห่งสัญญานี้ ผู้ว่าจ้างจะโอนเงินเข้าบัญชีเงินฝากธนาคารของ 
</br>ที่ปรึกษา ชื่อธนาคาร {result.ContractBankName} สาขา {result.ContractBankBranch} ชื่อบัญชี {result.ContractBankAccountName} เลขที่บัญชี {result.ContractBankAccountNumber} 
</br>ทั้งนี้ ที่ปรึกษาตกลงเป็นผู้รับภาระเงินค่าธรรมเนียมหรือค่าบริการอื่นใดเกี่ยวกับการโอน รวมทั้งค่าใช้จ่าย
</br>อื่นใด(ถ้ามี)ที่ธนาคารเรียกเก็บและยินยอมให้มีการหักเงินดังกล่าวจากจำนวนเงินโอนในงวดนั้นๆ 
</br>(ความในวรรคนี้ใช้สำหรับกรณีที่หน่วยงานของรัฐจะจ่ายเงินตรงให้แก่ที่ปรึกษา (ระบบ Direct Payment) โดยการโอนเงินเข้าบัญชีเงินฝากธนาคารของที่ปรึกษา ตามแนวทางที่กระทรวงการคลังหรือหน่วยงานของรัฐเจ้าของงบประมาณเป็นผู้กำหนด แล้วแต่กรณี)
    </p>
    <p class='tab3 t-16'>
        ค่าใช้จ่ายส่วนที่เบิกคืนได้(ถ้ามี)ผู้ว่าจ้างจะจ่ายคืนให้แก่ที่ปรึกษาสำหรับค่าใช้จ่ายซึ่งที่ปรึกษา
</br>ได้ใช้จ่ายไปตามความเป็นจริงตามเงื่อนไขที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก .....
    </p>
    <p class='tab3 t-16'><b>(๘)ข้อ ๔ เงินค่าจ้างล่วงหน้า</b></p>
    <p class='tab3 t-16'>
        ผู้ว่าจ้างตกลงจ่ายเงินค่าจ้างล่วงหน้าให้แก่ที่ปรึกษา เป็นจำนวนเงิน {result.PrepaidAmount?.ToString("N2") ?? "0.00"}
 บาท 
</br>({ToThaiNumberText(result.PrepaidAmount)}) ซึ่งเท่ากับร้อยละ {result.PrepaidPercents} ของค่าจ้าง
</br>ตามสัญญา
    </p>
    <p class='tab3 t-16'>
        เงินค่าจ้างล่วงหน้าดังกล่าวจะจ่ายให้ภายหลังจากที่ที่ปรึกษาได้วางหลักประกันการรับเงินค่า
</br>จ้างล่วงหน้าเป็น {result.GuaranteeType} (หนังสือค้ำประกันหรือหนังสือค้ำประกันอิเล็กทรอนิกส์ของธนาคารภายในประเทศ {result.PrepaidBankName} เต็มตามจำนวนเงินค่าจ้างล่วงหน้านั้นให้แก่ผู้ว่าจ้าง ที่ปรึกษาจะต้องออกใบเสร็จรับเงินค่า
</br>จ้างล่วงหน้าตามแบบที่ผู้ว่าจ้างกำหนดให้และที่ปรึกษาตกลงที่จะกระทำตามเงื่อนไขอันเกี่ยวกับการใช้จ่าย
</br>และการใช้คืนเงินค่าจ้างล่วงหน้านั้น ดังต่อไปนี้
    </p>
    <p class='tab3 t-16'>
        ๔.๑.ที่ปรึกษาจะใช้เงินค่าจ้างล่วงหน้านั้นเพื่อเป็นค่าใช้จ่ายในการปฏิบัติงานตามสัญญาเท่านั้นหากที่ปรึกษาใช้จ่ายเงินค่าจ้างล่วงหน้าหรือส่วนใดส่วนหนึ่งของเงินค่าจ้างล่วงหน้านั้นในทางอื่นผู้ว่าจ้างอาจจะเรียกเงินค่าจ้างล่วงหน้านั้นคืนจากที่ปรึกษาหรือบังคับเอาจากหลักประกันการรับเงินค่าจ้างล่วงหน้าได้ทันที
    </p>
    <p class='tab3 t-16'>
        ๔.๒.เมื่อผู้ว่าจ้างเรียกร้อง ที่ปรึกษาต้องแสดงหลักฐานการใช้จ่ายเงินค่าจ้างล่วงหน้า เพื่อ
</br>พิสูจน์ว่าได้เป็นไปตามข้อ ๔.๑ ภายในกำหนด ๑๕ (สิบห้า)วัน นับถัดจากวันได้รับแจ้งจากผู้ว่าจ้าง หากที่
</br>ปรึกษาไม่อาจแสดงหลักฐานดังกล่าว ภายในกำหนด ๑๕ (สิบห้า)วัน ผู้ว่าจ้างอาจเรียกเงินค่าจ้างล่วงหน้า
</br>คืนจากที่ปรึกษาหรือบังคับเอาจากหลักประกันการรับเงินค่าจ้างล่วงหน้าได้ทันที
    </p>
    <p class='tab3 t-16'>
        ๔.๓ ในการจ่ายเงินค่าจ้างให้แก่ที่ปรึกษาตามข้อ ๓ ผู้ว่าจ้างจะหักชดใช้คืนเงินค่าจ้างล่วงหน้า
</br>ในแต่ละงวดไว้จำนวนร้อยละ {result.PrepaidDeductPercent} ของจำนวนเงินค่าจ้างในแต่ละงวดจนกว่าจำนวนเงินที่หักไว้จะครบตาม
</br>จำนวนเงินที่หักค่าจ้างล่วงหน้าที่ที่ปรึกษาได้รับไปแล้ว ยกเว้นค่าจ้างงวดสุดท้ายจะหักไว้เป็นจำนวนเท่า

</br>กับจำนวนเงินค่าจ้างล่วงหน้าที่เหลือทั้งหมด
    </p>
    <p class='tab3 t-16'>
        ๔.๔ เงินจำนวนใดๆ ก็ตามที่ที่ปรึกษาจะต้องจ่ายให้แก่ผู้ว่าจ้างเพื่อชำระหนี้หรือเพื่อชดใช้
</br>ความรับผิดต่างๆ ตามสัญญา ผู้ว่าจ้างจะหักเอาจากเงินค่าจ้างงวดที่จะจ่ายให้แก่ที่ปรึกษาก่อนที่จะหัก
</br>ชดใช้คืนเงินค่าจ้างล่วงหน้า
    </p>
    <p class='tab3 t-16'>
        ๔.๕ ในกรณีที่มีการบอกเลิกสัญญา หากเงินค่าจ้างล่วงหน้าที่เหลือเกินกว่าจำนวนเงินที่ 
</br>ที่ปรึกษาจะได้รับหลังจากหักชดใช้ในกรณีอื่นแล้ว ที่ปรึกษาจะต้องจ่ายคืนเงินจำนวนที่เหลือนั้นให้แก่ผู้ว่าจ้าง 
</br>ภายใน ๗ (เจ็ด)วัน นับถัดจากวันได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้าง
    </p>
    <p class='tab3 t-16'>
        ๔.๖ ผู้ว่าจ้างจะคืนหลักประกันเงินค่าจ้างล่วงหน้าให้แก่ที่ปรึกษาต่อเมื่อผู้ว่าจ้างได้หักเงิน
</br>ค่าจ้างไว้จนครบจำนวนเงินค่าจ้างล่วงหน้าตามข้อ ๔.๓ แล้ว
    </p>
   <p class='tab3 t-16'><b>ข้อ ๕ ความรับผิดชอบของที่ปรึกษา</b></p>
    <p class='tab3 t-16'>
        ๕.๑ ที่ปรึกษาจะต้องส่งมอบผลงานตามรูปแบบและวิธีการ {result.SendWorkMethod} จำนวน {result.WorkAmount} ชุด ให้แก่ผู้ว่าจ้างตามที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก1
    </p>
    <p class='tab3 t-16'>
        ๕.๒ ในกรณีที่ผลงานของที่ปรึกษาบกพร่องหรือไม่เป็นไปตามข้อกำหนดและเงื่อนไข
</br>ตามสัญญาหรือมิได้ดำเนินการให้ถูกต้องตามหลักวิชาการ หรือวิชาชีพ {result.RelateExpertise} และ/หรือบทบัญญัติแห่ง
</br>กฎหมายที่เกี่ยวข้องที่ปรึกษาต้องรีบทำการแก้ไขให้เป็นที่เรียบร้อย โดยไม่คิดค่าจ้าง ค่าเสียหาย หรือ
</br>ค่าใช้จ่ายใดๆ จากผู้ว่าจ้างอีก ถ้าที่ปรึกษาหลีกเลี่ยงหรือไม่รีบจัดการแก้ไขให้เป็นที่เรียบร้อยในกำหนด
</br>เวลาที่ผู้ว่าจ้างแจ้งเป็นหนังสือผู้ว่าจ้างมีสิทธิจ้างที่ปรึกษารายอื่นทำการแทน โดยที่ปรึกษาจะต้องรับผิด
</br>ชอบจ่ายเงินค่าจ้างในการนี้แทนผู้ว่าจ้างโดยสิ้นเชิง
    </p>
    <p class='tab3 t-16'>
        ถ้ามีความเสียหายเกิดขึ้นจากงานตามสัญญานี้ไม่ว่าจะเนื่องมาจากการที่ที่ปรึกษา ได้ปฏิบัติ
</br>งานไม่ถูกต้องตามหลักวิชาการ หรือวิชาชีพ{result.RelateExpertise}และ/หรือบทบัญญัติแห่งกฎหมายที่เกี่ยวข้อง หรือเหตุใด
</br>ที่ปรึกษาจะต้องทำการแก้ไขความเสียหายดังกล่าว ภายในเวลาที่ผู้ว่าจ้างกำหนดให้ ถ้าที่ปรึกษา ไม่สามารถ
</br>แก้ไขได้ ที่ปรึกษาจะต้องชดใช้ค่าเสียหายที่เกิดขึ้นแก่ผู้ว่าจ้างโดยสิ้นเชิง ซึ่งรวมทั้งความเสียหายที่เกิดขึ้น
</br>โดยตรง และโดยส่วนที่เกี่ยวเนื่องกับความเสียหายที่เกิดขึ้นจากงานตามสัญญานี้ด้วย
    </p>
    <p class='tab3 t-16'>
        การที่ผู้ว่าจ้างได้ให้การรับรองหรือความเห็นชอบหรือความยินยอมใดๆ ในการปฏิบัติงาน
</br>หรือผลงานของที่ปรึกษาหรือการชำระเงินค่าจ้างตามสัญญาแก่ที่ปรึกษา ไม่เป็นการปลดเปลื้องพันธะและ 
</br>ความรับผิดชอบใดๆ ของที่ปรึกษาตามสัญญานี้
    </p>
    <p class='tab3 t-16'>
        ๕.๓ บุคลากรหลักของที่ปรึกษา ต้องมีระยะเวลาปฏิบัติงานตามสัญญานี้ไม่ซ้ำซ้อนกับงานใน
</br>โครงการอื่นๆ ของที่ปรึกษาที่ดำเนินการในช่วงเวลาเดียวกัน หากผู้ว่าจ้างพบว่าบุคลากรหลักไม่ว่าคนหนึ่ง
</br>คนใดหรือหลายคนปฏิบัติงานซ้ำซ้อนกับงานในโครงการอื่นๆ ไม่ว่าจะพบในระหว่างปฏิบัติงานตามสัญญา
</br>หรือในภายหลัง ผู้ว่าจ้างมีสิทธิบอกเลิกสัญญา และ/หรือเรียกค่าเสียหายจากที่ปรึกษาหรือปรับลดค่าจ้างได้
    </p>
    <p class='tab3 t-16'><b>ข้อ ๖ การระงับการทำงานชั่วคราวและการบอกเลิกสัญญา</b></p>
    <p class='tab3 t-16'>
        ๖.๑ ผู้ว่าจ้างมีสิทธิบอกเลิกสัญญาในกรณีดังต่อไปนี้
    </p>
    <p class='tab3 t-16'>
        (ก)หากผู้ว่าจ้างเห็นว่าที่ปรึกษามิได้ปฏิบัติงานด้วยความชำนาญหรือด้วยความเอาใจใส่ใน
</br>วิชาชีพของที่ปรึกษาเท่าที่พึงคาดหมายได้จากที่ปรึกษาในระดับเดียวกัน หรือมิได้ปฏิบัติตามสัญญาข้อใด
</br>ข้อหนึ่งในกรณีเช่นนี้ผู้ว่าจ้างจะบอกกล่าวให้ที่ปรึกษาทราบถึงเหตุผลที่จะบอกเลิกสัญญา ถ้าที่ปรึกษาผู้ว่าจ้าง
</br>มิได้ดำเนินการแก้ไขให้ผู้ว่าจ้างพอใจภายในระยะเวลา {result.FixDaysAfterNoti} วัน นับถัดจากวันที่ได้รับคำบอกกล่าว มีสิทธิบอก
</br>เลิกสัญญาโดยการส่งคำบอกกล่าวแก่ที่ปรึกษา เมื่อที่ปรึกษาได้รับหนังสือบอกกล่าวนั้นแล้วที่ปรึกษาต้องหยุด 
</br>ปฏิบัติงานทันที และดำเนินการทุกวิถีทางเพื่อลดค่าใช้จ่ายใดๆ ที่อาจมีในระหว่างการหยุดปฏิบัติงานนั้นให้
</br>น้อยที่สุด
    </p>
    <p class='tab3 t-16'>
        (ข)หากที่ปรึกษามิได้ลงมือทำงานภายในกำหนดเวลา หรือไม่สามารถทำงานให้แล้วเสร็จตาม
</br>กำหนดเวลา หรือมีเหตุให้ผู้ว่าจ้างเชื่อได้ว่าที่ปรึกษาไม่สามารถทำงานให้แล้วเสร็จภายในกำหนดเวลา หรือ 
</br>ล่วงเลยกำหนดเวลาแล้วเสร็จไปแล้ว หรือตกเป็นผู้ล้มละลาย ผู้ว่าจ้างมีสิทธิบอกเลิกสัญญาได้ทันที
    </p>
    <p class='tab3 t-16'>
        การบอกเลิกสัญญาตามข้อ ๖.๑ ผู้ว่าจ้างมีสิทธิริบหรือบังคับจากหลักประกันเงินค่าจ้างล่วง
</br>หน้า หลักประกันการปฏิบัติตามสัญญาและเงินประกันผลงานทั้งหมดหรือแต่บางส่วน และมีสิทธิเรียก
</br>ค่าเสียหายอื่น (ถ้ามี)จากที่ปรึกษาด้วย
    </p>
    <p class='tab3 t-16'>
        ๖.๒ ผู้ว่าจ้างอาจมีหนังสือบอกกล่าวให้ที่ปรึกษาทราบล่วงหน้าเมื่อใดก็ได้ว่าผู้ว่าจ้างมีเจตนา
</br>ที่จะระงับการทำงานของที่ปรึกษาไว้ชั่วคราวไม่ว่าทั้งหมดหรือแต่บางส่วน หรือจะบอกเลิกสัญญา ในกรณีที่
</br>ผู้ว่าจ้าง จะบอกเลิกสัญญา การบอกเลิกสัญญาดังกล่าวจะมีผลในเวลาไม่น้อยกว่า {result.NotiDaysAfterTerminate} วัน นับถัดจากวันที่
</br>ที่ปรึกษาได้รับหนังสือบอกกล่าวนั้น หรืออาจเร็วกว่าหรือช้ากว่ากำหนดเวลานั้นก็ได้แล้วแต่คู่สัญญาจะทำความ
</br>ตกลงกัน เมื่อที่ปรึกษาได้รับหนังสือบอกกล่าวนั้นแล้ว ที่ปรึกษาต้องหยุดปฏิบัติงานทันที ทั้งนี้ ที่ปรึกษา
</br>ไม่มีสิทธิได้รับค่าจ้างในระหว่างระงับการทำงานไว้ชั่วคราว และที่ปรึกษาจะต้องดำเนินการทุกวิถีทางเพื่อ
</br>ลดค่าใช้จ่ายใดๆ ที่อาจมีในระหว่างการหยุดปฏิบัติงานนั้นให้น้อยที่สุด
    </p>
    <p class='tab3 t-16'>
        กรณีที่มีการระงับการทำงานชั่วคราวตามข้อ ๖.๒ ผู้ว่าจ้างจะจ่ายเงินเป็นค่าใช้จ่ายเท่าที่
</br>จำเป็นให้แก่ที่ปรึกษาตามที่ผู้ว่าจ้างเห็นสมควร
    </p>
    <p class='tab3 t-16'>
        กรณีที่มีการเลิกสัญญาตามข้อ ๖.๒ ผู้ว่าจ้างจะชำระค่าจ้างตามส่วนที่เป็นธรรมและเหมาะสม
</br>ที่กำหนดในเอกสารแนบท้ายสัญญาผนวก ......ให้แก่ที่ปรึกษา โดยคำนวณตั้งแต่วันเริ่มปฏิบัติงานจนถึงวันบอก
</br>เลิกสัญญา นอกจากนี้ผู้ว่าจ้างจะคืนหลักประกันการปฏิบัติตามสัญญาหรือเงินประกันผลงาน รวมทั้งเงิน
</br>ชดเชยค่าเดินทางและเงินค่าใช้จ่ายที่ได้ทดรองจ่ายไปตามสมควรและตามความเป็นจริง ซึ่งผู้ว่าจ้างยังมิ
</br>ได้ชำระให้แก่ที่ปรึกษาด้วย อย่างไรก็ตาม เงินชดเชยและเงินที่ได้ชำระไปแล้วทั้งหมดจะต้องไม่เกิน
</br>ราคาค่าจ้างตามข้อ ๓
    </p>
    <p class='tab3 t-16'><b>ข้อ ๗ สิทธิและหน้าที่ของที่ปรึกษา</b></p>
    <p class='tab3 t-16'>
        ๗.๑ ที่ปรึกษาจะต้องใช้ความชำนาญ ความระมัดระวัง และความขยันหมั่นเพียร ในการ
</br>ปฏิบัติงานตามสัญญาอย่างมีประสิทธิภาพ และจะต้องปฏิบัติหน้าที่ตามความรับผิดชอบให้สำเร็จลุล่วง เป็นไป
</br>ตามมาตรฐานของวิชาชีพที่ยอมรับนับถือกันโดยทั่วไป
    </p>
    <p class='tab3 t-16'>
        ๗.๒ ค่าจ้างซึ่งผู้ว่าจ้างจะชำระแก่ที่ปรึกษาตามข้อ ๓ เป็นค่าตอบแทนเพียงอย่างเดียวเท่านั้น
</br>ซึ่งที่ปรึกษาจะได้รับเกี่ยวกับการปฏิบัติงานตามสัญญานี้ ที่ปรึกษาจะต้องไม่รับค่านายหน้าทางการค้า ส่วนลด 
</br>เบี้ยเลี้ยง เงินช่วยเหลือหรือผลประโยชน์ใดๆ ไม่ว่าโดยตรงหรือโดยอ้อม หรือสิ่งตอบแทนใดๆในส่วน
</br>ที่เกี่ยวข้องกับสัญญานี้ หรือที่เกี่ยวกับการปฏิบัติหน้าที่ตามสัญญานี้
    </p>
    <p class='tab3 t-16'>
        ๗.๓ ที่ปรึกษาจะต้องไม่มีผลประโยชน์ใดๆ ไม่ว่าโดยตรงหรือโดยอ้อมในเงินค่าสิทธิ 
</br>เงินบำเหน็จ หรือค่านายหน้าใดๆ ที่เกี่ยวกับการนำวัสดุสิ่งของหรือกรรมวิธีใดๆ ที่มีทะเบียนสิทธิบัตร
</br>หรือได้รับการคุ้มครองทางทรัพย์สินทางปัญญาหรือตามกฎหมายอื่นใดมาใช้เพื่อวัตถุประสงค์ของสัญญานี้
</br>เว้นแต่คู่สัญญา ทั้งสองฝ่ายจะได้ตกลงกันเป็นหนังสือว่าที่ปรึกษาอาจจะได้ผลประโยชน์หรือเงินเช่นว่านั้นได้
    </p>
    <p class='tab3 t-16'>
       ๗.๔ บรรดางานและเอกสารที่ที่ปรึกษาได้จัดทำขึ้นเกี่ยวกับสัญญานี้ให้ถือเป็นความลับและ
</br>ให้ตกเป็นกรรมสิทธิ์ของผู้ว่าจ้าง ที่ปรึกษาจะต้องส่งมอบบรรดางานและเอกสารดังกล่าวให้แก่ผู้ว่าจ้าง
</br>เมื่อสิ้นสุดสัญญานี้ ที่ปรึกษาอาจเก็บสำเนาเอกสารไว้กับตนได้แต่ต้องไม่นำข้อความในเอกสารนั้นไปใช้ใน
</br>กิจการอื่นที่ ไม่เกี่ยวกับงานโดยไม่ได้รับความยินยอมล่วงหน้าเป็นหนังสือจากผู้ว่าจ้างก่อน
    </p>
    <p class='tab3 t-16'>
        ๗.๕ ผู้ว่าจ้างเป็นเจ้าของลิขสิทธิ์หรือสิทธิในทรัพย์สินทางปัญญา รวมถึงสิทธิใดๆ ในผลงานที่
</br>ที่ปรึกษาได้ปฏิบัติงานตามสัญญานี้แต่เพียงฝ่ายเดียว และที่ปรึกษาจะนำผลงาน และ/หรือรายละเอียดของ
</br>งานตามสัญญานี้ ไม่ว่าทั้งหมดหรือบางส่วนไปใช้ หรือเผยแพร่ในกิจการอื่น นอกเหนือจากที่ได้ระบุ
</br>ไว้ในสัญญานี้ไม่ได้ เว้นแต่ได้รับอนุญาตเป็นหนังสือจากผู้ว่าจ้างก่อน
    </p>
    <p class='tab3 t-16'>
        ๗.๖ บรรดาเครื่องมือ เครื่องใช้ และวัสดุอุปกรณ์ทั้งหลาย ซึ่งผู้ว่าจ้างได้จัดให้ที่ปรึกษาใช้หรือ
</br>ซึ่งที่ปรึกษาซื้อมาด้วยทุนทรัพย์ของผู้ว่าจ้าง หรือซึ่งผู้ว่าจ้างเป็นผู้จ่ายชดใช้คืนให้ ถือว่าเป็นกรรมสิทธิ์ของ
</br>ผู้ว่าจ้างและต้องทำข้อความและเครื่องหมายที่แสดงว่าเป็นของผู้ว่าจ้างไว้ที่ทรัพย์สินดังกล่าวด้วย ทั้งนี้
</br>ที่ปรึกษาต้องใช้เครื่องมือเครื่องใช้และวัสดุอุปกรณ์ดังกล่าวอย่างเหมาะสมตามระเบียบของผู้ว่าจ้าง
</br>หรือของทางราชการ ที่เกี่ยวข้องเพื่อกิจการที่เกี่ยวกับการจ้างที่ปรึกษาเท่านั้น
    </p>
    <p class='tab3 t-16'>
        เมื่อที่ปรึกษาทำงานเสร็จหรือมีการเลิกสัญญา ที่ปรึกษาจะต้องทำบัญชีแสดงรายการเครื่อง
</br>มือเครื่องใช้และวัสดุอุปกรณ์ทั้งหลายข้างต้นที่ยังคงเหลืออยู่ และจัดการโยกย้ายไปเก็บรักษาตามคำสั่งผู้ว่าจ้าง
</br>โดยพลัน ที่ปรึกษาต้องดูแลเครื่องมือเครื่องใช้และวัสดุอุปกรณ์ดังกล่าวอย่างเหมาะสมตลอดเวลาที่ครอบครอง
</br>และต้องคืนเครื่องมือเครื่องใช้และวัสดุอุปกรณ์ดังกล่าวให้ครบถ้วนในสภาพดีตามความเหมาะสม 
</br>แต่ไม่ต้องรับผิดชอบสำหรับความเสื่อมสภาพจากการใช้งานตามปกติ
    </p>
    <p class='tab3 t-16'>
       ๗.๗ ที่ปรึกษาจะจัดให้มีบุคลากรที่มีความรู้และความชำนาญงานมาปฏิบัติงานให้เหมาะสมกับ
</br>สภาพการปฏิบัติงานตามสัญญานี้และให้สอดคล้องกับขอบเขตของงานของที่ปรึกษาตามที่ปรากฏ ในเอกสาร
</br>แนบท้ายสัญญาผนวก …..การเปลี่ยนแปลงบุคลากรดังกล่าวจะต้องได้รับความยินยอมเป็นหนังสือจากผู้ว่าจ้าง
</br>ก่อน
    </p>
    <p class='tab3 t-16'>
        ๗.๘ ในกรณีที่ผู้ว่าจ้างพิจารณาเห็นว่า การดำเนินงานของบุคลากรที่ที่ปรึกษาจัดหามาจะก่อ
</br>ให้เกิดความเสียหายแก่งานตามสัญญานี้ ไม่ว่าในกรณีใดก็ตามผู้ว่าจ้างมีสิทธิที่จะให้ที่ปรึกษาเปลี่ยนบุคลากร
</br>บางคนหรือทั้งหมดนั้นได้ และที่ปรึกษาต้องดำเนินการตามความประสงค์ของผู้ว่าจ้างโดยเร็ว
    </p>
    <p class='tab3 t-16'>
        การเปลี่ยนบุคลากรตามความในวรรคก่อน ที่ปรึกษาจะต้องเสนอรายชื่อบุคลากรที่จะ
</br>ปฏิบัติงานแทนนั้น ต่อผู้ว่าจ้างเพื่อพิจารณาให้ความเห็นชอบก่อน
    </p>
    <p class='tab3 t-16'><b>ข้อ ๘ ความรับผิดชอบของที่ปรึกษาต่อบุคคลภายนอก</b></p>
    <p class='tab3 t-16'>
        ๘.๑ ที่ปรึกษาจะต้องชดใช้ค่าเสียหายให้แก่ผู้ว่าจ้าง และป้องกันมิให้ผู้ว่าจ้างต้องรับผิดชอบใน
</br>บรรดาสิทธิเรียกร้อง ค่าเสียหาย ค่าใช้จ่าย หรือราคา รวมตลอดถึงการเรียกร้องโดยบุคคลภายนอกอันเกิดจาก
</br>ความผิดพลาดหรือการละเว้นไม่กระทำการของที่ปรึกษา หรือของลูกจ้างของที่ปรึกษา
    </p>
    <p class='tab3 t-16'>
        ๘.๒ ที่ปรึกษาจะต้องรับผิดชอบต่อการละเมิดบทบัญญัติแห่งกฎหมาย หรือการละเมิดลิขสิทธิ์ 
</br>หรือสิทธิในทรัพย์สินทางปัญญาอื่น รวมถึงสิทธิใดๆ ต่อบุคคลภายนอกเนื่องจากการปฏิบัติงานตามสัญญานี้ 
</br>โดยสิ้นเชิง
    </p>
    <p class='tab3 t-16'>
        (๙)๘.๓ ที่ปรึกษาจะต้องจัดการประกันภัยกับบริษัทประกันภัยที่ผู้ว่าจ้างเห็นชอบเพื่อความ
</br>รับผิดต่อบุคคลภายนอก และเพื่อความสูญหายหรือเสียหายในทรัพย์สินซึ่งผู้ว่าจ้างเป็นผู้จัดหาให้หรือสั่งซื้อ
</br>โดยทุนทรัพย์ของผู้ว่าจ้าง เพื่อให้ที่ปรึกษาไว้ใช้ในการปฏิบัติงานตามสัญญานี้ โดยที่ปรึกษาเป็นผู้ออกค่า
</br>ใช้จ่ายในการประกันภัยเอง ทั้งนี้ เว้นแต่จะมีการตกลงกันไว้เป็นอย่างอื่นในสัญญานี้
    </p>
    <p class='tab3 t-16'><b>ข้อ ๙ พันธะหน้าที่ของผู้ว่าจ้าง</b></p>
    <p class='tab3 t-16'>
        ผู้ว่าจ้างจะมอบข้อมูลและสถิติต่างๆ ที่เกี่ยวข้องซึ่งผู้ว่าจ้างมีอยู่ให้แก่ที่ปรึกษาโดยไม่คิดมูลค่า
</br>และภายในเวลาอันควร
    </p>
    <p class='tab3 t-16'>
        ในกรณีที่ที่ปรึกษาร้องขอความช่วยเหลือ ผู้ว่าจ้างจะพิจารณาให้ความช่วยเหลืออำนวยความ
</br>สะดวกตามสมควร ทั้งนี้ เพื่อให้การปฏิบัติงานของที่ปรึกษาตามสัญญานี้ลุล่วงไปได้ด้วยดี
    </p>
    <p class='tab3 t-16'><b>ข้อ ๑๐ ค่าปรับ</b></p>
    <p class='tab3 t-16'>
        ในกรณีที่ที่ปรึกษาไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่งด้วยเหตุใดๆ ก็ตาม จนเป็นเหตุให้เกิดค่า
</br>ปรับ ค่าเสียหาย หรือค่าใช้จ่ายแก่ผู้ว่าจ้าง ที่ปรึกษาต้องชดใช้ค่าปรับ ค่าเสียหายหรือค่าใช้จ่ายดังกล่าว
</br>ให้แก่ผู้ว่าจ้างโดยสิ้นเชิงภายในกำหนด {result.FinePerDays} วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้างหากที่ปรึกษา
</br>ไม่ชดใช้ให้ถูกต้องครบถ้วนภายในระยะเวลาดังกล่าว ให้ผู้ว่าจ้างมีสิทธิที่จะหักเอาจากจำนวนเงินค่าจ้างของ
</br>ที่ปรึกษาที่ต้องชำระ หรือบังคับจากหลักประกันการปฏิบัติตามสัญญาหรือเงินประกันผลงานได้ทันที
    </p>
    <p class='tab3 t-16'>
        หากค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายที่บังคับจากเงินค่าจ้างที่ต้องชำระ หลักประกัน 
</br>การปฏิบัติตามสัญญา และเงินประกันผลงานแล้วยังไม่เพียงพอ ที่ปรึกษายินยอมชำระส่วนที่เหลือที่ยังขาดอยู่
</br>จนครบถ้วนตามจำนวนค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายนั้น ภายในกำหนด {result.OutstandingPeriodDays}วันนับถัดจากวันที่ได้รับแจ้ง
</br>เป็นหนังสือจากผู้ว่าจ้าง
    </p>
    <p class='tab3 t-16'>
        หากมีเงินค่าจ้างตามสัญญาที่หักไว้จ่ายเป็นค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายแล้วยัง
</br>เหลืออยู่อีกเท่าใด ผู้ว่าจ้างจะคืนให้แก่ที่ปรึกษาทั้งหมด
    </p>
    <p class='tab3 t-16'><b>(๑๑)ข้อ ๑๒ (ก)เงินประกันผลงาน</b></p>
    <p class='tab3 t-16'>
        ในการจ่ายเงินให้แก่ที่ปรึกษาแต่ละงวด ผู้ว่าจ้างจะหักเงินจำนวนร้อยละ {result.RetentionRatePercent} ของเงินที่ต้อง
</br>จ่ายในงวดนั้นเพื่อเป็นประกันผลงาน หรือที่ปรึกษาอาจนำหนังสือค้ำประกันของธนาคารหรือหนังสือค้ำ
</br>ประกันอิเล็กทรอนิกส์ของธนาคารภายในประเทศซึ่งมีอายุการค้ำประกันตลอดอายุสัญญามามอบให้ผู้ว่า 
</br>จ้างทั้งนี้เพื่อเป็นหลักประกันแทนก็ได้
    </p>
    <p class='tab3 t-16'>
        ผู้ว่าจ้างจะคืนเงินประกันผลงาน และ/หรือหนังสือค้ำประกันของธนาคารดังกล่าวตามวรรค
</br>หนึ่งโดยไม่มีดอกเบี้ยให้แก่ที่ปรึกษาพร้อมกับการจ่ายเงินค่าจ้างงวดสุดท้าย
    </p>
    <p class='tab3 t-16'><b>ข้อ ๑๒ (ข)หลักประกันการปฏิบัติตามสัญญา</b></p>
    <p class='tab3 t-16'>
        ในขณะทำสัญญานี้ที่ปรึกษาได้นำหลักประกันเป็น{result.GuaranteeType}เป็นจำนวนเงิน {result.GuaranteeAmount} บาท 
</br>({ToThaiNumberText(result.GuaranteeAmount)})ซึ่งเท่ากับร้อยละ {result.GuaranteePercent} ของราคา
</br>ค่าจ้างตามสัญญา มามอบไว้แก่ผู้ว่าจ้างเพื่อเป็นหลักประกันการปฏิบัติตามสัญญานี้
    </p>
    <p class='tab3 t-16'>
        (๑๕)กรณีที่ปรึกษาใช้หนังสือค้ำประกันมาเป็นหลักประกันการปฏิบัติตามสัญญา หนังสือค้ำ
</br>ประกันดังกล่าวจะต้องออกโดยธนาคารที่ประกอบกิจการในประเทศไทย หรือโดยบริษัทเงินทุนหรือบริษัทเงิน
</br>ทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตาม
</br>ประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบ 
</br>โดยอนุโลมให้ใช้ตามตัวอย่างหนังสือค้ำประกันของธนาคารที่คณะกรรมการนโยบายกำหนด หรืออาจเป็น
</br>หนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้ และจะต้องมีอายุการค้ำประกัน
</br>ตลอดไปจนกว่าที่ปรึกษาพ้นข้อผูกพันตามสัญญานี้
    </p>
    <p class='tab3 t-16'>
        หลักประกันจะต้องมีอายุครอบคลุมความรับผิดทั้งปวงของที่ปรึกษาตลอดอายุสัญญา ถ้าหลัก
</br>ประกันที่ที่ปรึกษานำมามอบให้ดังกล่าวลดลงหรือเสื่อมค่าลง หรือมีอายุไม่ครอบคลุมถึงความรับผิดของ 
</br>ที่ปรึกษาตลอดอายุสัญญา ไม่ว่าด้วยเหตุใดๆ ก็ตาม รวมถึงกรณีที่ปรึกษาส่งมอบงานล่าช้าเป็นเหตุให้ระยะเวลา 
</br>แล้วเสร็จตามสัญญาเปลี่ยนแปลงไป ที่ปรึกษาต้องหาหลักประกันใหม่หรือหลักประกันเพิ่มเติมให้มีจำนวน
</br>ครบถ้วน ตามวรรคหนึ่งมามอบให้แก่ผู้ว่าจ้างภายใน {result.NewGuaranteeDays} วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้าง
    </p>
    <p class='tab3 t-16'>
        หลักประกันที่ผู้รับจ้างนำมามอบไว้ตามข้อนี้ ผู้ว่าจ้างจะคืนให้แก่ที่ปรึกษาโดยไม่มีดอกเบี้ย 
</br>เมื่อที่ปรึกษาพ้นจากข้อผูกพันและความรับผิดทั้งปวงตามสัญญานี้แล้ว
    </p>

    <p class='tab3 t-16'><b>ข้อ ๑๓ การจ้างช่วง</b></p>
    <p class='tab3 t-16'>
        ที่ปรึกษาจะต้องไม่เอางานทั้งหมดหรือแต่บางส่วนแห่งสัญญานี้ไปจ้างช่วงอีกทอดหนึ่ง เว้นแต่
</br>การจ้างช่วงงานแต่บางส่วนที่ได้รับอนุญาตเป็นหนังสือจากผู้ว่าจ้างก่อน การที่ผู้ว่าจ้างได้อนุญาตให้จ้างช่วงงาน
</br>แต่บางส่วนดังกล่าวนั้น ไม่เป็นเหตุให้ที่ปรึกษาหลุดพ้นจากความรับผิดหรือพันธะหน้าที่ตามสัญญานี้และที่
</br>ปรึกษาจะยังคงต้องรับผิดในความผิดและความประมาทเลินเล่อของผู้รับช่วงงาน หรือของตัวแทนหรือลูกจ้าง
</br>ของผู้รับช่วงงานนั้นทุกประการ

    </p>
 <p class='tab3 t-16'>
กรณีที่ปรึกษาไปจ้างช่วงงานแต่บางส่วนโดยฝ่าฝืนความในวรรคหนึ่ง ที่ปรึกษาต้องชำระค่า
</br>ปรับให้แก่ผู้ว่าจ้างเป็นจำนวนเงินในอัตราร้อยละ {result.SubcontractPenaltyPercent} ของวงเงินของงาน ที่จ้างช่วงตามสัญญา ทั้งนี้ 
</br>ไม่ตัดสิทธิผู้ว่าจ้างในการบอกเลิกสัญญา

</p>
  <p class='tab3 t-16'>
        การงดหรือลดค่าปรับ หรือขยายกำหนดเวลาทำงานตามวรรคหนึ่ง อยู่ในดุลพินิจของผู้ว่าจ้าง
</br>ที่จะพิจารณาตามที่เห็นสมควร
    </p>
    <p class='tab3 t-16'><b>ข้อ ๑๕ การงดหรือลดค่าปรับ หรือขยายเวลาปฏิบัติงานตามสัญญา</b></p>
    <p class='tab3 t-16'>
        ในกรณีที่มีเหตุเกิดจากความผิดหรือความบกพร่องของผู้ว่าจ้างหรือเหตุสุดวิสัย หรือเกิดจาก
</br>พฤติการณ์อันหนึ่งอันใดที่ที่ปรึกษาไม่ต้องรับผิดตามกฎหมาย หรือเหตุอื่นตามที่กำหนดในกฎกระทรวง ซึ่ง
</br>ออกตามความในกฎหมายว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ ทำให้ที่ปรึกษาไม่สามารถทำงาน
</br>ให้แล้วเสร็จตามเงื่อนไขและกำหนดเวลาแห่งสัญญานี้ได้ ที่ปรึกษาจะต้องแจ้งเหตุหรือพฤติการณ์ดังกล่าว 
</br>พร้อมหลักฐานเป็นหนังสือให้ผู้ว่าจ้างทราบ เพื่อของดหรือลดค่าปรับหรือขยายเวลาทำงานออกไปภายใน
</br> ๑๕ (สิบห้า)วัน นับถัดจากวันที่เหตุนั้นสิ้นสุดลง หรือตามที่กำหนดในกฎกระทรวงดังกล่าว แล้วแต่กรณี
    </p>
    <p class='tab3 t-16'>
        ถ้าที่ปรึกษาไม่ปฏิบัติให้เป็นไปตามความในวรรคหนึ่ง ให้ถือว่าที่ปรึกษาได้สละสิทธิเรียกร้อง 
</br>ในการที่จะของดหรือลดค่าปรับ หรือขยายเวลาทำงานออกไปโดยไม่มีเงื่อนไขใดๆ ทั้งสิ้นเว้นแต่กรณีเหตุ
</br>เกิดจากความผิดหรือความบกพร่องของฝ่ายผู้ว่าจ้างซึ่งมีหลักฐานชัดแจ้งหรือผู้ว่าจ้างทราบดีอยู่แล้วตั้งแต่ต้น
    </p>
    <p class='tab3 t-16'>
        การงดหรือลดค่าปรับ หรือขยายกำหนดเวลาทำงานตามวรรคหนึ่ง อยู่ในดุลพินิจของผู้ว่า
</br>จ้างที่จะพิจารณาตามที่เห็นสมควร
    </p>


{signatoryHtml}
  
    <!-- Add next page and appendix as needed -->

<div style='page-break-before: always;'></div>

    <div class='text-center t-22'><b>วิธีปฏิบัติเกี่ยวกับสัญญาจ้างผู้เชี่ยวชาญรายบุคคลหรือจ้างบริษัทที่ปรึกษา</b></div>

</br>

      <p class='tab2 t-16'>(๑)ให้ระบุเลขที่สัญญาในปีงบประมาณหนึ่งๆ ตามลำดับ</p>
    <p class='tab2 t-16'>(๒)ให้ระบุชื่อของหน่วยงานของรัฐที่เป็นนิติบุคคล เช่น กรม ก.หรือรัฐวิสาหกิจ ข.เป็นต้น</p>
    <p class='tab2 t-16'>(๓)ให้ระบุชื่อและตำแหน่งของหัวหน้าหน่วยงานของรัฐที่เป็นนิติบุคคลนั้น หรือผู้ที่ได้รับมอบอำนาจ เช่น นาย ก.อธิบดีกรม………...… หรือ นาย ข.ผู้ได้รับมอบอำนาจจากอธิบดีกรม………......………..</p>
    <p class='tab2 t-16'>(๔)ให้ระบุชื่อผู้ให้เช่า</p>
   <p class='tab3 t-16'>ก.กรณีนิติบุคคล เช่น ห้างหุ้นส่วนสามัญจดทะเบียน ห้างหุ้นส่วนจำกัด บริษัทจำกัด</p>
   <p class='tab3 t-16'>ข.กรณีบุคคลธรรมดา ให้ระบุชื่อและที่อยู่</p>
    <p class='tab2 t-16'>(๕)เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
    <p class='tab2 t-16'>(๖)เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
    <p class='tab2 t-16'>(๗)หน่วยงานของรัฐอาจกำหนดเงื่อนไขการจ่ายค่าเช่าให้แตกต่างไปจากแบบสัญญาที่กำหนดได้ตามความเหมาะสมและจำเป็นและไม่ทำให้หน่วยงานของรัฐเสียเปรียบ หากหน่วยงานของรัฐเห็นว่าจะมีปัญหาในทางเสียเปรียบหรือไม่รัดกุมพอ ก็ให้ส่งร่างสัญญานั้นไปให้สำนักงานอัยการสูงสุดพิจารณาให้ความเห็นชอบก่อน</p>
    <p class='tab2 t-16'>(๘)เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
    <p class='tab2 t-16'>(๙)ชื่อสถานที่หน่วยงานของรัฐ</p>
    <p class='tab2 t-16'>(๑๐)ให้พิจารณาถึงความจำเป็นและเหมาะสมของการใช้งาน</p>
    <p class='tab2 t-16'>(๑๑)อัตราค่าปรับตามสัญญาข้อ ๙ ให้กำหนดเป็นรายวันในอัตราระหว่างร้อยละ ๐.๐๑ – ๐.๒๐ ตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ.๒๕๖๗ ข้อ ๑๖๒ ส่วนกรณีจะปรับร้อยละเท่าใด ให้อยู่ในดุลยพินิจของหน่วยงานของรัฐผู้เช่าที่จะพิจารณา โดยคำนึงถึงราคาและลักษณะของพัสดุที่เช่า ซึ่งอาจมีผลกระทบต่อการที่ผู้ให้เช่าจะหลีกเลี่ยงไม่ปฏิบัติตามสัญญา แต่ทั้งนี้การที่จะกำหนดค่าปรับเป็นร้อยละเท่าใดจะต้องกำหนดไว้ในเอกสารเชิญชวนด้วย</p>
    <p class='tab2 t-16'>(๑๒)“หลักประกัน” หมายถึง หลักประกันที่ผู้ให้เช่านำมามอบไว้แก่หน่วยงานของรัฐ เมื่อลงนามในสัญญา เพื่อเป็นการประกันความเสียหายที่อาจจะเกิดขึ้นจากการปฏิบัติตามสัญญา ดังนี้</p>
   <p class='tab3 t-16'>(๑)เงินสด</p>
   <p class='tab3 t-16'>(๒)เช็คหรือดราฟท์ ที่ธนาคารเซ็นสั่งจ่าย ซึ่งเป็นเช็คหรือดราฟท์ลงวันที่ที่ใช้เช็คหรือดราฟท์นั้นชำระต่อเจ้าหน้าที่ หรือก่อนวันนั้นไม่เกิน ๓ วันทำการ</p>
   <p class='tab3 t-16'>(๓)หนังสือคํ้าประกันของธนาคารภายในประเทศตามตัวอย่างที่คณะกรรมการนโยบายกําหนด โดยอาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้</p>
   <p class='tab3 t-16'>(๔)หนังสือค้ำประกันของบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบ โดยอนุโลมให้ใช้ตามตัวอย่างหนังสือค้ำประกันของธนาคารที่คณะกรรมการนโยบายกำหนด</p>
   <p class='tab3 t-16'>(๕)พันธบัตรรัฐบาลไทย</p>
    <p class='tab2 t-16'>(๑๓)ให้กำหนดจำนวนเงินหลักประกันการปฏิบัติตามสัญญาตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ.๒๕๖๗ ข้อ ๑๖๘</p>
    <p class='tab2 t-16'>(๑๔)เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
    <p class='tab2 t-16'>(๑๕)กำหนดระยะเวลาตามความเหมาะสม เช่น 3 เดือน</p>
   <p class='tab3 t-16'>(๑๖)อัตราค่าปรับตามสัญญาข้อ ๑๒ ให้กำหนดเป็นรายวันในอัตราระหว่างร้อยละ ๐.๐๑ – ๐.๒๐ ตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ.๒๕๖๗ ข้อ ๑๖๒ ส่วนกรณีจะปรับร้อยละเท่าใด ให้อยู่ในดุลยพินิจของหน่วยงานของรัฐผู้เช่าที่จะพิจารณา โดยคำนึงถึงราคาและลักษณะของพัสดุที่เช่า ซึ่งอาจมีผลกระทบต่อการที่ผู้ให้เช่าจะหลีกเลี่ยงไม่ปฏิบัติตามสัญญา แต่ทั้งนี้การที่จะกำหนดค่าปรับเป็นร้อยละเท่าใดจะต้องกำหนดไว้ในเอกสารเชิญชวนด้วย</p>



</div>
";


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
         body {{
            font-size: 22px;
            font-family: 'THSarabunNew', Arial, sans-serif;
            word-break: break-word; 
         
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
        .tab1 {{ text-indent: 48px;     }}
        .tab2 {{ text-indent: 96px;    }}
        .tab3 {{ text-indent: 144px;    }}
        .tab4 {{ text-indent: 192px;   }}
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
        .table {{ width: 100%; border-collapse: collapse; margin-top: 20px;  }}
        .table th, .table td {{ border: 1px solid #000; padding: 8px; }}

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
</html>
";

        //        if (_pdfConverter == null)
        //            throw new Exception("PDF service is not available.");

        //        var doc = new HtmlToPdfDocument()
        //        {
        //            GlobalSettings = {
        //    PaperSize = PaperKind.A4,
        //    Orientation = DinkToPdf.Orientation.Portrait,
        //    Margins = new MarginSettings
        //    {
        //        Top = 20,
        //        Bottom = 20,
        //        Left = 20,
        //        Right = 20
        //    }
        //},
        //            Objects = {
        //    new ObjectSettings() {
        //        HtmlContent = html,
        //        FooterSettings = new FooterSettings
        //        {
        //            FontName = "THSarabunNew",
        //            FontSize = 6,
        //            Line = false,
        //            Center = "[page] / [toPage]"
        //        }
        //    }
        //}
        //        };

        //        var pdfBytes = _pdfConverter.Convert(doc);
                return html;

            }
        }
        catch (Exception ex)
        {
            throw new Exception("SLA 308-60 data not found.");
        }


    }
    #endregion 4.1.1.2.14.สัญญาจ้างที่ปรึกษา

}
