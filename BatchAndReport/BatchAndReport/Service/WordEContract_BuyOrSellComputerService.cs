using BatchAndReport.DAO;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Layout.Element;
using Microsoft.IdentityModel.Tokens;
using System.Threading.Tasks;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;


public class WordEContract_BuyOrSellComputerService
{
    private readonly WordServiceSetting _w;
    Econtract_Report_CPADAO _eContractReportDAO;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    private readonly EContractDAO _eContractDAO;
    public WordEContract_BuyOrSellComputerService(WordServiceSetting ws
        , Econtract_Report_CPADAO eContractReportDAO
         , IConverter pdfConverter
            , EContractDAO eContractDAO
        )
    {
        _w = ws;
        _eContractReportDAO = eContractReportDAO;
        _pdfConverter = pdfConverter;
        _eContractDAO = eContractDAO;
    }
# region 4.1.1.2.9.สัญญาเช่าคอมพิวเตอร์ 
    public async Task<byte[]> OnGetWordContact_BuyOrSellComputerService(string id)
    {
        var result = await _eContractReportDAO.GetCPAAsync(id);
        if (result == null)
        {
            throw new Exception("CPA data not found.");
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
                body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("สัญญาซื้อขายคอมพิวเตอร์", "000000", "36"));
                // 2. Document title and subtitle
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.RightParagraph("สัญญาเลขที่ "+result.CPAContractNumber??"xxxxxx"+""));


                // With this:
                string datestring = CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now);
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สัญญาฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.)\r\n"
                    + "ตำบล/แขวง ทุ่งสองห้อง อำเภอ/เขต หลักสี่\r\n" +
                "จังหวัด กรุงเทพ เมื่อ" + datestring + "\r\n" +
                "ระหว่าง " + result.Contract_Organization + "\r\n" +
                "โดย " + result.SignatoryName + "\r\n" +
                "ซึ่งต่อไปในสัญญานี้เรียกว่า “ผู้ซื้อ” ฝ่ายหนึ่ง กับ…" + result.ContractorName + "" , null, "32"));

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
                    " ผู้ถือบัตรประจำตัวประชาชนเลขที่ "+result.CitizenId+" ดังปรากฏตามสำเนาบัตรประจำตัวประชาชนแนบท้ายสัญญานี้) ซึ่งต่อไปใน\r\n" +
                    "สัญญานี้เรียกว่า “ผู้ให้เช่า” อีกฝ่ายหนึ่ง", null, "32"));

                }

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("คู่สัญญาได้ตกลงกันมีข้อความดังต่อไปนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ1 คำนิยาม", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ซื้อตกลงซื้อและผู้ขายตกลงขายและติดตั้งเครื่องคอมพิวเตอร์ อุปกรณ์การประมวลผล " +
                    "ระบบคอมพิวเตอร์ ซึ่งเป็นผลิตภัณฑ์ของ "+result.Computer_Model+"" +
                    "ซึ่งต่อไปในสัญญานี้เรียกว่า “คอมพิวเตอร์” ตามรายละเอียดเอกสารแนบท้ายสัญญาผนวก ๑ รวมเป็นราคาคอมพิวเตอร์และค่าติดตั้งทั้งสิ้น "+result.TotalAmount+" บาท ("+CommonDAO.NumberToThaiText(result.TotalAmount??0) + ")" +
                    "ซึ่งได้รวมภาษีมูลค่าเพิ่ม จำนวน "+result.VatAmount+" บาท ("+CommonDAO.NumberToThaiText(result.VatAmount ?? 0) + ") ตลอดจนภาษีอากรอื่นๆ และค่าใช้จ่ายทั้งปวงด้วยแล้ว", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้ขายประสงค์จะนำคอมพิวเตอร์และอุปกรณ์รายการใดแตกต่างไปจากรายละเอียดที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก๑ มาติดตั้งให้ผู้ซื้อ ผู้ขายจะต้องได้รับความเห็นชอบ" +
                    "เป็นหนังสือจากผู้ซื้อก่อน และคอมพิวเตอร์ที่จะนำมาติดตั้งดังกล่าวนั้นจะต้องมีคุณสมบัติไม่ต่ำกว่าที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก" +
                    "๑ และ ๒ ทั้งนี้ จะต้องไม่คิดราคาเพิ่มจากผู้ซื้อไม่ว่าในกรณีใด", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๒ การรับรองคุณภาพ", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ขายรับรองว่าคอมพิวเตอร์ที่ขายให้ตามสัญญานี้เป็นของแท้ ของใหม่ ไม่ใช่เครื่องที่ใช้งานแล้วนำมาปรับปรุงสภาพขึ้นใหม่และมีคุณสมบัติไม่ต่ำกว่าที่กำหนดไว้ตามรายละเอียด และคุณลักษณะเฉพาะของคอมพิวเตอร์ที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก ๒", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๓ เอกสารอันเป็นส่วนหนึ่งของสัญญา", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เอกสารแนบท้ายสัญญาดังต่อไปนี้ ให้ถือเป็นส่วนหนึ่งของสัญญานี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๓.๑ ผนวก ๑............(รายการคอมพิวเตอร์ที่ซื้อขาย)................จำนวน……..(.……….) หน้า", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๓.๒ ผนวก ๒...............(รายการคุณลักษณะเฉพาะ)..................จำนวน……..(.……….) หน้า", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๓.๓ผนวก ๓..(รายละเอียดการทดสอบการใช้งานคอมพิวเตอร์)..จำนวน……..(.……….) หน้า", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๓.๔ ผนวก ๔........(การกำหนดตัวถ่วงของคอมพิวเตอร์)...........จำนวน……..(.……….) หน้า", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๓.๕ ผนวก ๕........(การอบรมวิชาการด้านคอมพิวเตอร์)..........จำนวน……..(.……….) หน้า", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๓.๖ ผนวก ๖.....(รายการเอกสารคู่มือการใช้คอมพิวเตอร์).......จำนวน……..(.……….) หน้า", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs(".........................................ฯลฯ....................................", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ความใดในเอกสารแนบท้ายสัญญาที่ขัดหรือแย้งกับข้อความในสัญญานี้ ให้ใช้ข้อความ" +
                    "ในสัญญานี้บังคับ และในกรณีที่เอกสารแนบท้ายสัญญาขัดแย้งกันเอง ผู้ขายจะต้องปฏิบัติตามคำวินิจฉัยของผู้ซื้อ คำวินิจฉัยของผู้ซื้อให้ถือเป็นที่สุด และผู้ขายไม่มีสิทธิเรียกร้องราคา ค่าเสียหาย หรือค่าใช้จ่ายใด ๆ เพิ่มเติมจากผู้ซื้อทั้งสิ้น", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๔ การส่งมอบและติดตั้ง", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ขายจะส่งมอบและติดตั้งคอมพิวเตอร์ที่ซื้อขายตามสัญญานี้ให้ถูกต้องและครบถ้วนตามที่กำหนดไว้ในข้อ ๑ ให้พร้อมที่จะใช้งานได้ตามรายละเอียดการทดสอบการใช้งานคอมพิวเตอร์ เอกสารแนบท้ายสัญญาผนวก๓" +
                    "ให้แก่ผู้ซื้อ ณ "+result.DeliveryLocation+" และส่งมอบให้แก่ผู้ซื้อภายใน "+result.DeliveryDateIn+"  วัน นับถัดจากวันลงนามในสัญญา", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ขายจะต้องแจ้งกำหนดเวลาติดตั้งแล้วเสร็จพร้อมที่จะใช้งานและส่งมอบคอมพิวเตอร์ได้" +
                    "โดยทำเป็นหนังสือยื่นต่อผู้ซื้อ ณ "+result.Contract_Sign_Address+" ในวันและเวลาทำการของผู้ซื้อก่อนวันกำหนดส่งมอบไม่น้อยกว่า "+result.NotiDaysBeforeDelivery+" วันทำการของผู้ซื้อ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ขายต้องออกแบบสถานที่ติดตั้งคอมพิวเตอร์รวมทั้งระบบอื่นๆ ที่เกี่ยวข้องตามมาตรฐานของผู้ขายและได้รับความเห็นชอบจากผู้ซื้อเป็นหนังสือ และผู้ขายต้องจัดหาเจ้าหน้าที่มาให้คำแนะนำและตรวจสอบความถูกต้องเหมาะสมของสถานที่ให้ทันต่อการติดตั้งคอมพิวเตอร์โดยไม่คิดค่าใช้จ่ายใด ๆ จากผู้ซื้อภายใน "+result.LocationPrepareDays+" วัน นับถัดจากวันลงนามในสัญญา", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๕ การตรวจรับ", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เมื่อผู้ซื้อได้ตรวจรับคอมพิวเตอร์ที่ส่งมอบและติดตั้งแล้วเห็นว่าถูกต้องครบถ้วนตามสัญญานี้แล้ว ผู้ซื้อจะออกหลักฐานการรับมอบไว้เป็นหนังสือ เพื่อผู้ขายนำมาใช้เป็นหลักฐานประกอบการขอรับเงินค่าคอมพิวเตอร์ ตามที่ระบุไว้ในข้อ ๖", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ถ้าผลของการตรวจรับปรากฏว่า คอมพิวเตอร์ที่ผู้ขายส่งมอบไม่ตรงตามข้อ ๑ หรือมีคุณสมบัติไม่ถูกต้องตามข้อ ๒ หรือใช้งานได้ไม่ครบถ้วนตามข้อ ๔ ผู้ซื้อทรงไว้ซึ่งสิทธิที่จะไม่รับคอมพิวเตอร์นั้น ในกรณีเช่นว่านี้ ผู้ขายต้องรีบนำคอมพิวเตอร์นั้นกลับคืนโดยเร็วที่สุดเท่าที่จะทำได้และนำคอมพิวเตอร์มาส่งมอบให้ใหม่ หรือต้องทำการแก้ไขให้ถูกต้องตามสัญญาด้วยค่าใช้จ่ายของผู้ขายเอง และระยะเวลาที่เสียไปเพราะเหตุดังกล่าวผู้ขายจะนำมาอ้างเป็นเหตุขอขยายเวลาทำการตามสัญญาหรือของดหรือลดค่าปรับไม่ได้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(8) ในกรณีที่ผู้ขายส่งมอบคอมพิวเตอร์ถูกต้องแต่ไม่ครบจำนวน หรือส่งมอบครบจำนวนแต่ไม่ถูกต้องทั้งหมด ผู้ซื้อจะตรวจรับเฉพาะส่วนที่ถูกต้อง โดยออกหลักฐานการตรวจรับเฉพาะส่วนนั้นก็ได้ (ความในวรรคสามนี้ จะไม่กำหนดไว้ในกรณีที่ผู้ซื้อต้องการคอมพิวเตอร์ทั้งหมดในคราวเดียวกัน หรือการซื้อคอมพิวเตอร์ที่ประกอบเป็นชุดหรือหน่วย ถ้าขาดส่วนประกอบอย่างหนึ่งอย่างใดไปแล้ว จะไม่สามารถใช้งานได้โดยสมบูรณ์)", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๖ การชำระเงิน", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(9) ผู้ซื้อจะชำระเงินค่าคอมพิวเตอร์ให้แก่ผู้ขาย เมื่อผู้ซื้อได้รับมอบคอมพิวเตอร์ตามข้อ ๕ ไว้โดยครบถ้วนแล้ว", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(10) ผู้ซื้อตกลงชำระเงินค่าคอมพิวเตอร์ให้แก่ผู้ขาย ดังนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๖.๑ เงินล่วงหน้า จำนวน "+result.AdvancePayment+" บาท ("+CommonDAO.NumberToThaiText(result.AdvancePayment??0) + ") " +
                    "จะจ่ายให้ภายใน \"+result.PaymentDueDays+\" วัน นับถัดจากวันลงนามในสัญญา ทั้งนี้ โดยผู้ขายจะต้อง" +
                    "นำหลักประกันการรับเงินล่วงหน้าเป็น. " + result.PaymentGuaranteeType + "(หนังสือค้ำประกันหรือหนังสือค้ำประกันอิเล็กทรอนิกส์" +
                    "ของธนาคารภายในประเทศหรือพันธบัตรรัฐบาลไทย)" + result.PaymentGuaranteeTypeOther + " เต็มตามจำนวนเงินล่วงที่ได้รับ " +
                    "มามอบให้แก่ผู้ซื้อเป็นหลักประกันการชำระคืนเงินล่วงหน้าก่อนการรับชำระเงินล่วงหน้านั้น และผู้ซื้อจะคืนหลักประกันการรับเงินล่วงหน้าให้แก่ผู้ขาย เมื่อผู้ซื้อจ่ายเงินที่เหลือตามข้อ ๖.๒ แล้ว เว้นแต่ในกรณีดังต่อไปนี้ " +
                    "ผู้ขายมีสิทธิขอคืนหลักประกันการรับเงินล่วงหน้าบางส่วนก่อนได้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๑) กรณีผู้ขายได้วางหลักประกันการรับเงินล่วงหน้าไว้ฉบับเดียว หากผู้ซื้อได้หักเงินล่วงหน้า" +
                    "ไปแล้ว ผู้ขายมีสิทธิขอคืนหลักประกันการรับเงินล่วงหน้าในส่วนที่ผู้ซื้อได้หักเงินล่วงหน้าไปแล้วนั้น โดยผู้ขายจะต้องนำหลักประกันการรับเงินล่วงหน้าฉบับใหม่ที่มีมูลค่าเท่ากับเงินล่วงหน้าที่เหลืออยู่มาวางให้แก่ผู้ซื้อ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๒) กรณีผู้ขายได้วางหลักประกันการรับเงินล่วงหน้าไว้หลายฉบับ ซึ่งแต่ละฉบับมีมูลค่าเท่ากับจำนวนเงินล่วงหน้าที่ผู้ซื้อจะต้องหักไว้ในแต่ละงวด หากผู้ซื้อได้หักเงินล่วงหน้าในงวดใดแล้วผู้ขายมีสิทธิขอคืนหลักประกันการรับเงินล่วงหน้าในงวดนั้นได้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๖.๒ เงินที่เหลือ จำนวน……......…….........บาท (…....…………...………………) จะจ่ายให้เมื่อผู้ซื้อได้รับคอมพิวเตอร์ตามข้อ ๕ ไว้โดยถูกต้องครบถ้วน และผู้ขายได้อบรมเจ้าหน้าที่ของผู้ซื้อตามข้อ ๑๐ เสร็จสิ้นแล้ว", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(11) การจ่ายเงินตามเงื่อนไขแห่งสัญญานี้ ผู้ซื้อจะโอนเงินเข้าบัญชีเงินฝากธนาคารของผู้ขาย ชื่อธนาคาร "+result.SaleBankName+" สาขา "+result.SaleBankBranch+" ชื่อบัญชี "+result.SaleBankAccountName+" เลขที่บัญชี"+result.SaleBankAccountNumber+" ทั้งนี้ ผู้ขายตกลงเป็นผู้รับภาระเงินค่าธรรมเนียม หรือค่าบริการอื่นใดเกี่ยวกับการโอน รวมทั้งค่าใช้จ่ายอื่นใด (ถ้ามี) ที่ธนาคารเรียกเก็บและยินยอมให้มีการหักเงินดังกล่าวจากจำนวนเงินโอนในงวดนั้น ๆ (ความในวรรคนี้ใช้สำหรับกรณีที่หน่วยงานของรัฐจะจ่ายเงินตรงให้แก่ผู้ขาย (ระบบ Direct Payment) โดยการโอนเงินเข้าบัญชีเงินฝากธนาคารของผู้ขายตามแนวทางที่กระทรวงการคลังหรือหน่วยงานของรัฐเจ้าของงบประมาณเป็นผู้กำหนด แล้วแต่กรณี)", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๗ การรับประกันความชำรุดบกพร่อง", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ขายตกลงรับประกันความชำรุดบกพร่องหรือขัดข้องของคอมพิวเตอร์และการติดตั้ง" +
                    "ตามสัญญานี้เป็นเวลา "+result.WarrantyPeriodYears+" ปี "+result.WarrantyPeriodMonths+" เดือน นับถัดจากวันที่" +
                    "ผู้ซื้อได้รับมอบคอมพิวเตอร์ทั้งหมดโดยถูกต้องครบถ้วนตามสัญญา ถ้าภายในระยะเวลาดังกล่าวคอมพิวเตอร์ชำรุด" +
                    "บกพร่องหรือขัดข้อง หรือใช้งานไม่ได้ทั้งหมดหรือแต่บางส่วน หรือเกิดความชำรุดบกพร่องหรือขัดข้องจากการติดตั้ง เว้นแต่ความชำรุดบกพร่องหรือขัดข้องดังกล่าว เกิดขึ้นจากความผิดของผู้ซื้อซึ่งไม่ได้เกิดขึ้นจากการใช้งานตามปกติ ผู้ขายจะต้องจัดการซ่อมแซมแก้ไขให้อยู่ในสภาพใช้การได้ดีดังเดิม โดยต้องเริ่มจัดการซ่อมแซมแก้ไขภายใน "+result.DaysToRepairAfterNoti+" วัน นับถัดจากวันที่ได้รับแจ้งจากผู้ซื้อโดยไม่คิดค่าใช้จ่ายใด ๆ จากผู้ซื้อทั้งสิ้น ถ้าผู้ขายไม่จัดการซ่อมแซมแก้ไขภายในกำหนดเวลาดังกล่าว ผู้ซื้อมีสิทธิที่จะทำการนั้นเองหรือจ้างผู้อื่นทำการนั้นแทนผู้ขาย โดยผู้ขายต้องออกค่าใช้จ่ายเองทั้งสิ้นแทนผู้ซื้อ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ขายมีหน้าที่บำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ให้อยู่ในสภาพใช้งานได้ดีอยู่เสมอตลอดระยะเวลาดังกล่าวในวรรคหนึ่งด้วยค่าใช้จ่ายของผู้ขาย โดยให้มีเวลาคอมพิวเตอร์ขัดข้องรวมตามเกณฑ์การคำนวณเวลาขัดข้องไม่เกินเดือนละ "+result.DaysToRepairAfterNoti+" ชั่วโมง หรือร้อยละ "+result.MaximumDownTimePercent+"(%) ของเวลาใช้งานทั้งหมดของคอมพิวเตอร์ของเดือนนั้น แล้วแต่ตัวเลขใดจะมากกว่ากัน มิฉะนั้นผู้ขายต้องยอมให้ผู้ซื้อคิดค่าปรับเป็นรายชั่วโมง ในอัตราร้อยละ "+result.PenaltyPerHourPercent+" ของราคาคอมพิวเตอร์ทั้งหมดตามสัญญานี้คิดเป็นเงิน "+result.PenaltyPerHour+" บาท  ("+CommonDAO.NumberToThaiText(result.TotalAmount??0) + ") ต่อชั่วโมง ในช่วงเวลาที่ไม่สามารถใช้คอมพิวเตอร์ได้ในส่วนที่เกินกว่ากำหนดเวลาขัดข้องข้างต้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เกณฑ์การคำนวณเวลาขัดข้องของคอมพิวเตอร์ตามวรรคสอง ให้เป็นดังนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("- กรณีที่คอมพิวเตอร์เกิดขัดข้องพร้อมกันหลายหน่วย ให้นับเวลาขัดข้องของหน่วยที่มีตัวถ่วงมากที่สุดเพียงหน่วยเดียว", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("- กรณีความเสียหายอันสืบเนื่องมาจากความขัดข้องของคอมพิวเตอร์แตกต่างกัน เวลาที่ใช้ในการคำนวณค่าปรับจะเท่ากับเวลาขัดข้องของคอมพิวเตอร์หน่วยนั้นคูณด้วยตัวถ่วงซึ่งมีค่าต่างๆ ตามเอกสาร แนบท้ายสัญญาผนวก ๔", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("การที่ผู้ซื้อทำการนั้นเอง หรือให้ผู้อื่นทำการนั้นแทนผู้ขาย ไม่ทำให้ผู้ขายหลุดพ้นจาก ความรับผิดตามสัญญา หากผู้ขายไม่ชดใช้ค่าใช้จ่ายหรือค่าเสียหายตามที่ผู้ซื้อเรียกร้องผู้ซื้อมีสิทธิบังคับจากหลักประกันการปฏิบัติตามสัญญาได้", null, "32"));



                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๘ หลักประกันการปฏิบัติตามสัญญา", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในขณะทำสัญญานี้ผู้ขายได้นำหลักประกันเป็น "+result.PerformanceGuarantee+" เป็นจำนวนเงิน "+result.GuaranteeAmount+" บาท  ("+CommonDAO.NumberToThaiText(result.GuaranteeAmount??0) + ") ซึ่งเท่ากับร้อยละ  "+result.GuaranteePercent+" ของราคาซื้อขายคอมพิวเตอร์ตามข้อ ๑ มามอบให้แก่ผู้ซื้อเพื่อเป็นหลักประกันการปฏิบัติตามสัญญานี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(16) กรณีผู้ขายใช้หนังสือค้ำประกันมาเป็นหลักประกันการปฏิบัติตามสัญญา หนังสือค้ำประกันดังกล่าวจะต้องออกโดยธนาคารที่ประกอบกิจการในประเทศไทย หรือโดยบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบตามแบบที่คณะกรรมการนโยบายการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐกำหนด หรืออาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้และจะต้องมีอายุการค้ำประกันตลอดไปจนกว่าผู้ขายพ้นข้อผูกพันตามสัญญานี้ ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หลักประกันที่ผู้ขายนำมามอบให้ตามวรรคหนึ่ง จะต้องมีอายุครอบคลุมความรับผิดทั้งปวงของผู้ขายตลอดอายุสัญญา ถ้าหลักประกันที่ผู้ขายนำมามอบให้ดังกล่าวลดลงหรือเสื่อมค่าลง หรือมีอายุไม่ครอบคลุมถึงความรับผิดของผู้ขายตลอดอายุสัญญา ไม่ว่าด้วยเหตุใดๆ ก็ตาม รวมถึงกรณีผู้ขายส่งมอบและติดตั้งคอมพิวเตอร์ล่าช้าเป็นเหตุให้ระยะเวลาแล้วเสร็จหรือวันครบกำหนดความรับผิดในความชำรุดบกพร่องตามสัญญาเปลี่ยนแปลงไป ไม่ว่าจะเกิดขึ้นคราวใด ผู้ขายต้องหาหลักประกันใหม่หรือหลักประกันเพิ่มเติมให้มีจำนวนครบถ้วนตามวรรคหนึ่ง มามอบให้แก่ผู้ซื้อภายใน "+result.NewGuaranteeDays+" วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ซื้อ ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หลักประกันที่ผู้ขายนำมามอบไว้ตามข้อนี้ ผู้ซื้อจะคืนให้แก่ผู้ขายโดยไม่มีดอกเบี้ยเมื่อผู้ขายพ้นจากข้อผูกพันและความรับผิดทั้งปวงตามสัญญานี้แล้ว", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๙ การโอนกรรมสิทธิ์", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("คู่สัญญาตกลงกันว่ากรรมสิทธิ์ในคอมพิวเตอร์ตามสัญญาจะโอนไปยังผู้ซื้อเมื่อผู้ซื้อได้รับมอบคอมพิวเตอร์ดังกล่าวตามข้อ ๕ แล้ว", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๐ การอบรม", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ขายต้องจัดอบรมวิชาการด้านคอมพิวเตอร์ให้แก่เจ้าหน้าที่ของผู้ซื้อจนสามารถใช้งานคอมพิวเตอร์ได้อย่างมีประสิทธิภาพ โดยต้องดำเนินการฝึกอบรมให้แล้วเสร็จภายใน "+result.TrainingPeriodDays+" วันนับถัดจากวันที่ผู้ซื้อได้รับมอบคอมพิวเตอร์ โดยไม่คิดค่าใช้จ่ายใด ๆ รายละเอียดของการฝึกอบรมให้เป็นไปตามเอกสารแนบท้ายสัญญาผนวก ๕", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ข้อ ๑๑ คู่มือการใช้คอมพิวเตอร์", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ขายต้องจัดหาและส่งมอบคู่มือการใช้คอมพิวเตอร์ตามสัญญานี้ ตามที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก ๖ จำนวน "+result.ComputerManualsCount+" ชุด ให้กับผู้ซื้อในวันที่ส่งมอบคอมพิวเตอร์ พร้อมทั้งปรับปรุงให้ทันสมัยเป็นปัจจุบันตลอดอายุสัญญานี้ ทั้งนี้ โดยไม่คิดเงินเพิ่มจากผู้ซื้อ", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๒ การรับประกันความเสียหาย", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่บุคคลภายนอกกล่าวอ้างหรือใช้สิทธิเรียกร้องใด ๆ ว่ามีการละเมิดลิขสิทธิ์หรือสิทธิบัตรหรือสิทธิใด ๆ เกี่ยวกับคอมพิวเตอร์ตามสัญญานี้ โดยผู้ซื้อมิได้แก้ไขดัดแปลงไปจากเดิม ผู้ขายจะต้องดำเนินการทั้งปวงเพื่อให้การกล่าวอ้างหรือการเรียกร้องดังกล่าวระงับสิ้นไปโดยเร็ว หากผู้ซื้อต้องรับผิดชดใช้ค่าเสียหายต่อบุคคลภายนอกเนื่องจากผลแห่งการละเมิดลิขสิทธิ์หรือสิทธิบัตรหรือสิทธิใด ๆ ดังกล่าว ผู้ขายต้องเป็นผู้ชำระค่าเสียหายและค่าใช้จ่ายรวมทั้งค่าฤชาธรรมเนียมและค่าทนายความแทนผู้ซื้อ ทั้งนี้ ผู้ซื้อจะแจ้งให้ผู้ขายทราบเป็นหนังสือในเมื่อได้มีการกล่าวอ้างหรือใช้สิทธิเรียกร้องดังกล่าวโดยไม่ชักช้า", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๓ การบอกเลิกสัญญา", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เมื่อครบกำหนดส่งมอบคอมพิวเตอร์ตามสัญญาแล้ว หากผู้ขายไม่ส่งมอบและติดตั้งคอมพิวเตอร์บางรายการหรือทั้งหมดให้แก่ผู้ซื้อภายในกำหนดเวลาดังกล่าว หรือส่งมอบคอมพิวเตอร์ไม่ตรงตามสัญญาหรือมีคุณสมบัติไม่ถูกต้องตามสัญญา หรือส่งมอบและติดตั้งแล้วเสร็จภายในกำหนดแต่ไม่สามารถใช้งานได้อย่างมีประสิทธิภาพ หรือใช้งานไม่ได้ครบถ้วนตามสัญญา หรือผู้ขายไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่ง ผู้ซื้อมีสิทธิบอกเลิกสัญญาทั้งหมดหรือแต่บางส่วนได้ การใช้สิทธิบอกเลิกสัญญานั้นไม่กระทบถึงสิทธิของผู้ซื้อที่จะเรียกร้องค่าเสียหายจากผู้ขาย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้ซื้อใช้สิทธิบอกเลิกสัญญา ผู้ซื้อมีสิทธิริบหรือบังคับจากหลักประกันตาม (17) (ข้อ ๖ และ) ข้อ ๘ เป็นจำนวนเงินทั้งหมดหรือแต่บางส่วนก็ได้ แล้วแต่ผู้ซื้อจะเห็นสมควร และถ้าผู้ซื้อจัดซื้อคอมพิวเตอร์รวมถึงการติดตั้งจากบุคคลอื่นเต็มจำนวนหรือเฉพาะจำนวนที่ขาดส่ง แล้วแต่กรณี ภายในกำหนด…"+result.TeminationNewMonths+" เดือน นับถัดจากวันบอกเลิกสัญญา ผู้ขายจะต้องชดใช้ราคาที่เพิ่มขึ้นจากราคาที่กำหนดไว้ในสัญญานี้ด้วย รวมทั้งค่าใช้จ่ายใด ๆ ที่ผู้ซื้อต้องใช้จ่ายในการจัดหาผู้ขายรายใหม่ดังกล่าวด้วย ในกรณีที่ผู้ขายได้ส่งมอบคอมพิวเตอร์ให้แก่ผู้ซื้อและผู้ซื้อบอกเลิกสัญญา ผู้ขายจะต้องนำคอมพิวเตอร์กลับคืนไป และทำสถานที่ที่รื้อถอนคอมพิวเตอร์ออกไปให้มีสภาพดังที่มีอยู่เดิมก่อนทำสัญญานี้ภายใน "+result.ReturnDaysIn+" วัน นับถัดจากวันที่ผู้ซื้อบอกเลิกสัญญา โดยผู้ขายเป็นผู้เสียค่าใช้จ่ายเองทั้งสิ้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(19)ถ้าผู้ขายไม่ยอมนำคอมพิวเตอร์กลับคืนไปภายในกำหนดเวลาดังกล่าวตามวรรคสาม ผู้ซื้อจะกำหนดเวลาให้ผู้ขายนำคอมพิวเตอร์กลับคืนไปอีกครั้งหนึ่ง หากพ้นกำหนดเวลาดังกล่าวแล้ว ผู้ขายยังไม่นำคอมพิวเตอร์กลับคืนไปอีก ผู้ซื้อมีสิทธินำคอมพิวเตอร์ออกขายทอดตลาด เงินที่ได้จากการขายทอดตลาด ผู้ขายยอมให้ผู้ซื้อหักเป็นค่าปรับและหักเป็นค่าใช้จ่าย และค่าเสียหายที่เกิดแก่ผู้ซื้อ ซึ่งรวมถึงค่าใช้จ่ายต่างๆ ที่ผู้ซื้อได้เสียไปในการดำเนินการขายทอดตลาดคอมพิวเตอร์ดังกล่าว ค่าใช้จ่ายในการทำสถานที่ที่รื้อถอนคอมพิวเตอร์ออกไปให้มีสภาพดังที่มีอยู่เดิมก่อนทำสัญญานี้ เงินที่เหลือจากการหักค่าปรับ ค่าใช้จ่าย และค่าเสียหายดังกล่าวแล้วผู้ซื้อจะคืนให้แก่ผู้ขาย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เมื่อผู้ซื้อบอกเลิกสัญญาแล้ว ผู้ซื้อไม่ต้องรับผิดชอบในความเสียหายใด ๆ ทั้งสิ้นอันเกิดแก่คอมพิวเตอร์ซึ่งอยู่ในความครอบครองของผู้ซื้อ ", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๔ ค่าปรับ", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้ซื้อยังไม่ใช้สิทธิบอกเลิกสัญญาตามข้อ ๑๓ ผู้ขายจะต้องชำระค่าปรับให้ผู้ซื้อเป็นรายวัน ในอัตราร้อยละ "+result.FinePerDaysPercent+" ของราคาคอมพิวเตอร์ที่ยังไม่ได้รับมอบ นับถัดจากวันครบกำหนดตามสัญญาจนถึงวันที่ผู้ขายได้นำคอมพิวเตอร์มาส่งมอบและติดตั้งให้แก่ผู้ซื้อจนถูกต้องครบถ้วนตามสัญญา", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("การคิดค่าปรับในกรณีที่คอมพิวเตอร์ที่ตกลงซื้อขายเป็นระบบ ถ้าผู้ขายส่งมอบเพียงบางส่วนหรือขาดส่วนประกอบส่วนหนึ่งส่วนใดไป หรือส่งมอบและติดตั้งทั้งหมดแล้วแต่ใช้งานไม่ได้ถูกต้องครบถ้วน ให้ถือว่ายังไม่ได้ส่งมอบคอมพิวเตอร์นั้นเลย และคิดค่าปรับจากราคาคอมพิวเตอร์ทั้งระบบ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในระหว่างที่ผู้ซื้อยังไม่ได้ใช้สิทธิบอกเลิกสัญญานั้น หากผู้ซื้อเห็นว่าผู้ขายไม่อาจปฏิบัติตามสัญญาต่อไปได้ ผู้ซื้อจะใช้สิทธิบอกเลิกสัญญา และริบหรือบังคับจากหลักประกันตาม (21) (ข้อ ๖ และ) ข้อ ๘ กับเรียกร้องให้ชดใช้ราคาที่เพิ่มขึ้นตามที่กำหนดไว้ในข้อ ๑๓ วรรคสอง ก็ได้ และถ้าผู้ซื้อได้แจ้งข้อเรียกร้องให้ชำระค่าปรับไปยังผู้ขายเมื่อครบกำหนดส่งมอบแล้ว ผู้ซื้อมีสิทธิที่จะปรับผู้ขายจนถึงวันบอกเลิกสัญญาได้อีกด้วย", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๕ การบังคับค่าปรับ ค่าเสียหาย และค่าใช้จ่าย", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้ขายไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่งด้วยเหตุใด ๆ ก็ตาม จนเป็นเหตุให้เกิดค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายแก่ผู้ซื้อ ผู้ขายต้องชดใช้ค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายดังกล่าวให้แก่ผู้ซื้อ   โดยสิ้นเชิงภายในกำหนด "+result.EnforcementOfFineDays+" วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ซื้อ หากผู้ขายไม่ชดใช้ให้ถูกต้องครบถ้วนภายในระยะเวลาดังกล่าวให้ผู้ซื้อมีสิทธิที่จะหักเอาจากค่าคอมพิวเตอร์ที่ต้องชำระ หรือบังคับจากหลักประกันการปฏิบัติตามสัญญาได้ทันที", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หากค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายที่บังคับจากค่าคอมพิวเตอร์ที่ต้องชำระหรือหลักประกันการปฏิบัติตามสัญญาแล้วยังไม่เพียงพอ ผู้ขายยินยอมชำระส่วนที่เหลือที่ยังขาดอยู่จนครบถ้วนตามจำนวนค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายนั้น ภายในกำหนด "+result.OutstandingPeriodDays+" วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ซื้อ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หากมีเงินค่าคอมพิวเตอร์ที่ซื้อขายตามสัญญาที่หักไว้จ่ายเป็นค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายแล้วยังเหลืออยู่อีกเท่าใด ผู้ซื้อจะคืนให้แก่ผู้ขายทั้งหมด", null, "32"));



                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๖ การงดหรือลดค่าปรับ หรือขยายเวลาในการปฏิบัติตามสัญญา", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่มีเหตุเกิดจากความผิดหรือความบกพร่องของผู้ซื้อ หรือเหตุสุดวิสัย หรือเกิดจากพฤติการณ์อันหนึ่งอันใดที่ผู้ขายไม่ต้องรับผิดตามกฎหมาย" +
                    "หรือเหตุอื่นตามที่กำหนดในกฎกระทรวง ซึ่งออกตามความในกฎหมายว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ ทำให้ผู้ขายไม่สามารถส่งมอบและติดตั้งคอมพิวเตอร์ " +
                    "หรือบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ ตามเงื่อนไขและกำหนดเวลาแห่งสัญญานี้ได้ ผู้ขายมีสิทธิของดหรือลดค่าปรับหรือขยายเวลาทำการตามสัญญา โดยจะต้องแจ้งเหตุหรือพฤติการณ์ดังกล่าวพร้อมหลักฐาน" +
                    "เป็นหนังสือให้ผู้ซื้อทราบภายใน ๑๕ (สิบห้า) วัน นับถัดจากวันที่เหตุนั้นสิ้นสุดลง หรือตามที่กำหนดในกฎกระทรวงดังกล่าว ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ถ้าผู้ขายไม่ปฏิบัติให้เป็นไปตามความในวรรคหนึ่ง ให้ถือว่าผู้ขายได้สละสิทธิเรียกร้องในการที่จะของดหรือลดค่าปรับหรือขยายเวลาทำการตามสัญญา โดยไม่มีเงื่อนไขใด ๆ ทั้งสิ้น เว้นแต่กรณีเหตุเกิดจากความผิดหรือความบกพร่องของฝ่ายผู้ซื้อซึ่งมีหลักฐานชัดแจ้งหรือผู้ซื้อทราบดีอยู่แล้วตั้งแต่ต้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("การงดหรือลดค่าปรับหรือขยายเวลาทำการตามสัญญาตามวรรคหนึ่ง อยู่ในดุลพินิจของ    ผู้ซื้อที่จะพิจารณาตามที่เห็นสมควร", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๗ การใช้เรือไทย", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ถ้าผู้ขายจะต้องสั่งหรือนำเข้าคอมพิวเตอร์มาจากต่างประเทศและต้องนำเข้ามาโดยทางเรือในเส้นทางเดินเรือที่มีเรือไทยเดินอยู่ และสามารถให้บริการรับขนได้ตามที่รัฐมนตรีว่าการกระทรวงคมนาคมประกาศกำหนด ผู้ขายต้องจัดการให้คอมพิวเตอร์บรรทุกโดยเรือไทยหรือเรือที่มีสิทธิเช่นเดียวกับเรือไทยจากต่างประเทศมายังประเทศไทย เว้นแต่จะได้รับอนุญาตจากกรมเจ้าท่าก่อนบรรทุกคอมพิวเตอร์ลงเรืออื่นที่มิใช่เรือไทยหรือเป็นของที่รัฐมนตรีว่าการกระทรวงคมนาคมประกาศยกเว้นให้บรรทุกโดยเรืออื่นได้ ทั้งนี้ ไม่ว่าการสั่งหรือนำเข้าคอมพิวเตอร์จากต่างประเทศจะเป็นแบบใด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในการส่งมอบคอมพิวเตอร์ให้แก่ผู้ซื้อ ถ้าเป็นกรณีตามวรรคหนึ่งผู้ขายจะต้องส่งมอบใบตราส่ง (Bill of Lading) หรือสำเนาใบตราส่งสำหรับคอมพิวเตอร์นั้น ซึ่งแสดงว่าได้บรรทุกมาโดยเรือไทย หรือเรือที่มีสิทธิเช่นเดียวกับเรือไทยให้แก่ผู้ซื้อพร้อมกับการส่งมอบคอมพิวเตอร์ด้วย", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่คอมพิวเตอร์ไม่ได้บรรทุกจากต่างประเทศมายังประเทศไทยโดยเรือไทยหรือเรือที่มีสิทธิเช่นเดียวกับเรือไทย ผู้ขายต้องส่งมอบหลักฐานซึ่งแสดงว่าได้รับอนุญาตจากกรมเจ้าท่าให้บรรทุกของโดยเรืออื่นได้ หรือหลักฐานซึ่งแสดงว่าได้ชำระค่าธรรมเนียมพิเศษเนื่องจากการไม่บรรทุกของโดยเรือไทยตามกฎหมายว่าด้วยการส่งเสริมการพาณิชยนาวีแล้วอย่างใดอย่างหนึ่งแก่ผู้ซื้อด้วย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้ขายไม่ส่งมอบหลักฐานอย่างใดอย่างหนึ่งดังกล่าวในวรรคสองและวรรคสามให้แก่ผู้ซื้อแต่จะขอส่งมอบคอมพิวเตอร์ให้ผู้ซื้อก่อนโดยยังไม่รับชำระเงินค่าคอมพิวเตอร์ ผู้ซื้อมีสิทธิรับคอมพิวเตอร์ไว้ก่อนและชำระเงินค่าคอมพิวเตอร์เมื่อผู้ขายได้ปฏิบัติถูกต้องครบถ้วนดังกล่าวแล้วได้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("สัญญานี้ทำขึ้นสองฉบับ มีข้อความถูกต้องตรงกัน คู่สัญญาได้อ่านและเข้าใจข้อความโดยละเอียดตลอดแล้ว จึงได้ลงลายมือชื่อพร้อมทั้งประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยานและคู่สัญญาต่างยึดถือไว้ฝ่ายละหนึ่งฉบับ", null, "32"));



                body.AppendChild(WordServiceSetting.EmptyParagraph());

                body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ "+result.OSMEP_Signer??"xxxxxxxxxx"+"ผู้ซื้อ"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(................................................................................)"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ "+result.Contract_Signer+" ผู้ขาย"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(................................................................................)"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ"+result.OSMEP_Witness?? "พยาน SME" + "พยาน"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(...............................................................................)"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ" + result.Contract_Witness ?? "พยานขาย" + "พยาน"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(...............................................................................)"));

                // next page
                body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("วิธีปฏิบัติเกี่ยวกับสัญญาซื้อขายคอมพิวเตอร์", "000000", "36"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(1) ให้ระบุเลขที่สัญญาในปีงบประมาณหนึ่งๆ ตามลำดับ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(2) ให้ระบุชื่อของหน่วยงานของรัฐที่เป็นนิติบุคคล เช่น กรม ก. หรือรัฐวิสาหกิจ ข. เป็นต้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(3) ให้ระบุชื่อและตำแหน่งของหัวหน้าหน่วยงานของรัฐที่เป็นนิติบุคคลนั้น หรือผู้ที่ได้รับมอบอำนาจ เช่น นาย ก. อธิบดีกรม………...… หรือ นาย ข. ผู้ได้รับมอบอำนาจจากอธิบดีกรม………......………..", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(4) ให้ระบุชื่อผู้รับจ้าง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ก. กรณีนิติบุคคล เช่น ห้างหุ้นส่วนสามัญจดทะเบียน ห้างหุ้นส่วนจำกัด บริษัทจำกัด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ข. กรณีบุคคลธรรมดา ให้ระบุชื่อและที่อยู่", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(5) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(6) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๗) ให้ระบุยี่ห้อคอมพิวเตอร์ รุ่นคอมพิวเตอร์", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๘) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(9) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(10) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(11) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(12) ระยะเวลารับประกันและระยะเวลาแก้ไขซ่อมแซมจะกำหนดเท่าใด แล้วแต่ลักษณะของสิ่งของที่ซื้อขายกัน โดยให้อยู่ในดุลพินิจของผู้ซื้อ ทั้งนี้ จะต้องประกาศให้ทราบในเอกสารประกวดราคาด้วย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(13) ให้กำหนดในอัตราระหว่างร้อยละ ๐.๐๒๕ – ๐.๐๓๕ ของราคาตามสัญญาต่อชั่วโมง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(14) “หลักประกัน” หมายถึง หลักประกันที่ผู้รับจ้างนำมามอบไว้แก่หน่วยงานของรัฐเมื่อลงนามในสัญญาเพื่อประกันความเสียหายที่อาจเกิดขึ้นจากการปฏิบัติตามสัญญา ดังนี้ ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๑)เงินสด ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๒)เช็คหรือดราฟท์ที่ธนาคารเซ็นสั่งจ่าย ซึ่งเป็นเช็คหรือดราฟท์ลงวันที่ที่ใช้เช็คหรือดราฟท์นั้นชำระต่อเจ้าหน้าที่ หรือก่อนวันนั้นไม่เกิน ๓ วันทำการ ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๓)หนังสือค้ำประกันของธนาคารภายในประเทศตามตัวอย่างที่คณะกรรมการนโยบายกำหนด โดยอาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๔)หนังสือค้ำประกันของบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบ โดยอนุโลมให้ใช้ตามตัวอย่างหนังสือ         ค้ำประกันของธนาคารที่คณะกรรมการนโยบายกำหนด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๕)พันธบัตรรัฐบาลไทย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(15) ให้กำหนดจำนวนเงินหลักประกันการปฏิบัติตามสัญญาตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ ข้อ ๑๖๘ ทั้งนี้ จะต้องประกาศให้ทราบในเอกสารประกวดราคาด้วย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(16) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(17) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(18) กำหนดเวลาที่ผู้ซื้อจะซื้อสิ่งของจากแหล่งอื่นเมื่อบอกเลิกสัญญาและมีสิทธิเรียกเงินในส่วนที่เพิ่มขึ้นจากราคาที่กำหนดไว้ในสัญญานั้น ให้อยู่ในดุลพินิจของผู้ซื้อโดยตกลงกับผู้ขาย ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(19) ความในวรรคนี้อาจแก้ไขเปลี่ยนแปลงได้ตามความเหมาะสมถ้าหน่วยงานของรัฐผู้ทำสัญญาสามารถกำหนดมาตรการอื่นใดในสัญญา หรือกำหนดทางปฏิบัติเพื่อแก้ปัญหาที่ผู้ขายไม่ยอมนำคอมพิวเตอร์กลับคืนไปได้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("20) อัตราค่าปรับตามสัญญาข้อ ๑๔ ให้กำหนดเป็นรายวันในอัตราตายตัวระหว่างร้อยละ ๐.๐๑ – ๐.๒๐ ของราคาพัสดุที่ยังไม่ได้ส่งมอบ ตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ ข้อ ๑๖๒ ส่วนกรณีจะปรับร้อยละเท่าใด ให้อยู่ในดุลพินิจของหน่วยงานของรัฐผู้ซื้อที่จะพิจารณา โดยคำนึงถึงราคาและลักษณะของพัสดุที่ซื้อ ซึ่งอาจมีผลกระทบต่อการที่ผู้ขายจะหลีกเลี่ยงไม่ปฏิบัติตามสัญญา แต่ทั้งนี้ การที่จะกำหนดค่าปรับเป็นร้อยละเท่าใดจะต้องกำหนดไว้ในเอกสารเชิญชวนด้วย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(21) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));

                body.AppendChild(WordServiceSetting.EmptyParagraph());




                WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);

            }
            stream.Position = 0;
            return stream.ToArray();
        }
 
    }
    public async Task<byte[]> OnGetWordContact_BuyOrSellComputerService_ToPDF(string id)
    {
        var result = await _eContractReportDAO.GetCPAAsync(id);
        if (result == null)
            throw new Exception("CPA data not found.");
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");
        var cssPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "css", "contract.css").Replace("\\", "/");
        var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
        string logoBase64 = "";
        if (System.IO.File.Exists(logoPath))
        {
            var bytes = System.IO.File.ReadAllBytes(logoPath);
            logoBase64 = Convert.ToBase64String(bytes);
        }

        var strDateTH = CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now);

        //เอกสารแนบท้ายสัญญาผนวก 1-6
        var listDocAtt = await _eContractDAO.GetRelatedDocumentsAsync(id, "CPA");
        var htmlDocAtt = "";
        if (listDocAtt != null)
        {
            for (int i = 0; i < listDocAtt.Count; i++)
            {
                var docItem = listDocAtt[i];
                htmlDocAtt += $"<div class='tab3 t-16'> {docItem.DocumentTitle} จำนวน {docItem.PageAmount} หน้า </div>";
                          
                          
            }
        }
        // Build HTML content
        var htmlContent = $@"
<div >
    <div class='text-center t-22'><b> แบบสัญญา</b></div>
    <div class='text-center t-22'><b> สัญญาซื้อขายคอมพิวเตอร์</b></div>
</br>
    <div class='text-right t-18'>สัญญาเลขที่ {(result.CPAContractNumber ?? "xxxxxx")}</div>
    <div class='tab3 t-16'>
        สัญญาฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.) ตำบล/แขวง ทุ่งสองห้อง อำเภอ/เขต หลักสี่ 
        จังหวัด กรุงเทพ เมื่อ {strDateTH} ระหว่าง {result.Contract_Organization} 
        โดย {result.SignatoryName} ซึ่งต่อไปในสัญญานี้เรียกว่า “ผู้ซื้อ” ฝ่ายหนึ่ง กับ…{result.ContractorName}
    </div>
    <div class='tab3 t-16'>
        {(result.ContractorType == "นิติบุคคล"
       ? $"ซึ่งจดทะเบียนเป็นนิติบุคคล ณ {result.ContractorName} มี " +
         $"สำนักงานใหญ่อยู่เลขที่ {result.ContractorAddressNo} ถนน {result.ContractorStreet} ตำบล/แขวง {result.ContractorSubDistrict} " +
         $"อำเภอ/เขต {result.ContractorDistrict} จังหวัด {result.ContractorProvince}<br/>โดย {result.ContractorSignatoryName} " +
         $"มีอำนาจลงนามผูกพันนิติบุคคลปรากฏตามหนังสือรับรองของสำนักงานทะเบียนหุ้นส่วนบริษัท …………… " +
         $"ลงวันที่ {CommonDAO.ToThaiDateString(result.ContractSignDate ?? DateTime.Now)} (5)(และหนังสือมอบอำนาจลง {CommonDAO.ToThaiDateString(result.ContractSignDate ?? DateTime.Now)}) แนบท้ายสัญญานี้"
       : $"(6)(ในกรณีที่ผู้รับจ้างเป็นบุคคลธรรมดาให้ใช้ข้อความว่า กับ {result.ContractorName} " +
         $"อยู่บ้านเลขที่ {result.ContractorAddressNo} ถนน {result.ContractorStreet} ตำบล/แขวง {result.ContractorSubDistrict} " +
         $"อำเภอ/เขต {result.ContractorDistrict} จังหวัด {result.ContractorProvince} " +
         $"ผู้ถือบัตรประจำตัวประชาชนเลขที่ {result.CitizenId} ดังปรากฏตามสำเนาบัตรประจำตัวประชาชนแนบท้ายสัญญานี้) ซึ่งต่อไปในสัญญานี้เรียกว่า “ผู้ให้เช่า” อีกฝ่ายหนึ่ง")}
    </div>
    <div class='tab3 t-16'>คู่สัญญาได้ตกลงกันมีข้อความดังต่อไปนี้</div>
    <div class='tab2 t-16'>ข้อ1 คำนิยาม</div>
    <div class='tab3 t-16'>
        ผู้ซื้อตกลงซื้อและผู้ขายตกลงขายและติดตั้งเครื่องคอมพิวเตอร์ อุปกรณ์การประมวลผล
        ระบบคอมพิวเตอร์ ซึ่งเป็นผลิตภัณฑ์ของ {result.Computer_Model}
        ซึ่งต่อไปในสัญญานี้เรียกว่า “คอมพิวเตอร์” ตามรายละเอียดเอกสารแนบท้ายสัญญาผนวก ๑ รวมเป็นราคาคอมพิวเตอร์และค่าติดตั้งทั้งสิ้น {result.TotalAmount} บาท ({CommonDAO.NumberToThaiText(result.TotalAmount ?? 0)})
        ซึ่งได้รวมภาษีมูลค่าเพิ่ม จำนวน {result.VatAmount} บาท ({CommonDAO.NumberToThaiText(result.VatAmount ?? 0)}) ตลอดจนภาษีอากรอื่นๆ และค่าใช้จ่ายทั้งปวงด้วยแล้ว
    </div>
    <div class='tab3 t-16'>
        ในกรณีที่ผู้ขายประสงค์จะนำคอมพิวเตอร์และอุปกรณ์รายการใดแตกต่างไปจากรายละเอียดที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก๑ มาติดตั้งให้ผู้ซื้อ ผู้ขายจะต้องได้รับความเห็นชอบเป็นหนังสือจากผู้ซื้อก่อน และคอมพิวเตอร์ที่จะนำมาติดตั้งดังกล่าวนั้นจะต้องมีคุณสมบัติไม่ต่ำกว่าที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก ๑ และ ๒ ทั้งนี้ จะต้องไม่คิดราคาเพิ่มจากผู้ซื้อไม่ว่าในกรณีใด
    </div>
    <div class='tab2 t-16'>ข้อ ๒ การรับรองคุณภาพ</div>
    <div class='tab3 t-16'>
        ผู้ขายรับรองว่าคอมพิวเตอร์ที่ขายให้ตามสัญญานี้เป็นของแท้ ของใหม่ ไม่ใช่เครื่องที่ใช้งานแล้วนำมาปรับปรุงสภาพขึ้นใหม่และมีคุณสมบัติไม่ต่ำกว่าที่กำหนดไว้ตามรายละเอียด และคุณลักษณะเฉพาะของคอมพิวเตอร์ที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก ๒
    </div>
    <div class='tab2 t-16'>ข้อ ๓ เอกสารอันเป็นส่วนหนึ่งของสัญญา</div>
    <div class='tab3 t-16'>
        เอกสารแนบท้ายสัญญาดังต่อไปนี้ ให้ถือเป็นส่วนหนึ่งของสัญญานี้<br/></div>
        {htmlDocAtt} 
    <div class='tab3 t-16'>
        ความใดในเอกสารแนบท้ายสัญญาที่ขัดหรือแย้งกับข้อความในสัญญานี้ ให้ใช้ข้อความในสัญญานี้บังคับ และในกรณีที่เอกสารแนบท้ายสัญญาขัดแย้งกันเอง ผู้ขายจะต้องปฏิบัติตามคำวินิจฉัยของผู้ซื้อ คำวินิจฉัยของผู้ซื้อให้ถือเป็นที่สุด และผู้ขายไม่มีสิทธิเรียกร้องราคา ค่าเสียหาย หรือค่าใช้จ่ายใด ๆ เพิ่มเติมจากผู้ซื้อทั้งสิ้น
    </div>
    <div class='tab2 t-16'>ข้อ ๔ การส่งมอบและติดตั้ง</div>
    <div class='tab3 t-16'>
        ผู้ขายจะส่งมอบและติดตั้งคอมพิวเตอร์ที่ซื้อขายตามสัญญานี้ให้ถูกต้องและครบถ้วนตามที่กำหนดไว้ในข้อ ๑ ให้พร้อมที่จะใช้งานได้ตามรายละเอียดการทดสอบการใช้งานคอมพิวเตอร์ เอกสารแนบท้ายสัญญาผนวก๓ ให้แก่ผู้ซื้อ ณ {result.DeliveryLocation} และส่งมอบให้แก่ผู้ซื้อภายใน {result.DeliveryDateIn} วัน นับถัดจากวันลงนามในสัญญา<br/>
        ผู้ขายจะต้องแจ้งกำหนดเวลาติดตั้งแล้วเสร็จพร้อมที่จะใช้งานและส่งมอบคอมพิวเตอร์ได้โดยทำเป็นหนังสือยื่นต่อผู้ซื้อ ณ {result.Contract_Sign_Address} ในวันและเวลาทำการของผู้ซื้อก่อนวันกำหนดส่งมอบไม่น้อยกว่า {result.NotiDaysBeforeDelivery} วันทำการของผู้ซื้อ<br/>
        ผู้ขายต้องออกแบบสถานที่ติดตั้งคอมพิวเตอร์รวมทั้งระบบอื่นๆ ที่เกี่ยวข้องตามมาตรฐานของผู้ขายและได้รับความเห็นชอบจากผู้ซื้อเป็นหนังสือ และผู้ขายต้องจัดหาเจ้าหน้าที่มาให้คำแนะนำและตรวจสอบความถูกต้องเหมาะสมของสถานที่ให้ทันต่อการติดตั้งคอมพิวเตอร์โดยไม่คิดค่าใช้จ่ายใด ๆ จากผู้ซื้อภายใน {result.LocationPrepareDays} วัน นับถัดจากวันลงนามในสัญญา
    </div>
    <!-- Continue with more sections as in your Word logic, e.g. ข้อ ๕ การตรวจรับ, ข้อ ๖ การชำระเงิน, ... -->
    <div class='tab2 t-16'>ข้อ ๕ การตรวจรับ</div>
    <div class='tab3 t-16'>
        เมื่อผู้ซื้อได้ตรวจรับคอมพิวเตอร์ที่ส่งมอบและติดตั้งแล้วเห็นว่าถูกต้องครบถ้วนตามสัญญานี้แล้ว ผู้ซื้อจะออกหลักฐานการรับมอบไว้เป็นหนังสือ เพื่อผู้ขายนำมาใช้เป็นหลักฐานประกอบการขอรับเงินค่าคอมพิวเตอร์ ตามที่ระบุไว้ในข้อ ๖<br/>
        ถ้าผลของการตรวจรับปรากฏว่า คอมพิวเตอร์ที่ผู้ขายส่งมอบไม่ตรงตามข้อ ๑ หรือมีคุณสมบัติไม่ถูกต้องตามข้อ ๒ หรือใช้งานได้ไม่ครบถ้วนตามข้อ ๔ ผู้ซื้อทรงไว้ซึ่งสิทธิที่จะไม่รับคอมพิวเตอร์นั้น ในกรณีเช่นว่านี้ ผู้ขายต้องรีบนำคอมพิวเตอร์นั้นกลับคืนโดยเร็วที่สุดเท่าที่จะทำได้และนำคอมพิวเตอร์มาส่งมอบให้ใหม่ หรือต้องทำการแก้ไขให้ถูกต้องตามสัญญาด้วยค่าใช้จ่ายของผู้ขายเอง และระยะเวลาที่เสียไปเพราะเหตุดังกล่าวผู้ขายจะนำมาอ้างเป็นเหตุขอขยายเวลาทำการตามสัญญาหรือของดหรือลดค่าปรับไม่ได้<br/>
        (8) ในกรณีที่ผู้ขายส่งมอบคอมพิวเตอร์ถูกต้องแต่ไม่ครบจำนวน หรือส่งมอบครบจำนวนแต่ไม่ถูกต้องทั้งหมด ผู้ซื้อจะตรวจรับเฉพาะส่วนที่ถูกต้อง โดยออกหลักฐานการตรวจรับเฉพาะส่วนนั้นก็ได้ (ความในวรรคสามนี้ จะไม่กำหนดไว้ในกรณีที่ผู้ซื้อต้องการคอมพิวเตอร์ทั้งหมดในคราวเดียวกัน หรือการซื้อคอมพิวเตอร์ที่ประกอบเป็นชุดหรือหน่วย ถ้าขาดส่วนประกอบอย่างหนึ่งอย่างใดไปแล้ว จะไม่สามารถใช้งานได้โดยสมบูรณ์)
    </div>
   <!-- ข้อ ๖ การชำระเงิน -->
<div class='tab2 t-16'>ข้อ ๖ การชำระเงิน</div>
<div class='tab3 t-16'>
    (9) ผู้ซื้อจะชำระเงินค่าคอมพิวเตอร์ให้แก่ผู้ขาย เมื่อผู้ซื้อได้รับมอบคอมพิวเตอร์ตามข้อ ๕ ไว้โดยครบถ้วนแล้ว<br/>
    (10) ผู้ซื้อตกลงชำระเงินค่าคอมพิวเตอร์ให้แก่ผู้ขาย ดังนี้<br/>
    ๖.๑ เงินล่วงหน้า จำนวน {result.AdvancePayment} บาท ({CommonDAO.NumberToThaiText(result.AdvancePayment ?? 0)}) จะจ่ายให้ภายใน {result.PaymentDueDays} วัน นับถัดจากวันลงนามในสัญญา ทั้งนี้ โดยผู้ขายจะต้องนำหลักประกันการรับเงินล่วงหน้าเป็น {result.PaymentGuaranteeType} {result.PaymentGuaranteeTypeOther} เต็มตามจำนวนเงินล่วงที่ได้รับ มามอบให้แก่ผู้ซื้อเป็นหลักประกันการชำระคืนเงินล่วงหน้าก่อนการรับชำระเงินล่วงหน้านั้น และผู้ซื้อจะคืนหลักประกันการรับเงินล่วงหน้าให้แก่ผู้ขาย เมื่อผู้ซื้อจ่ายเงินที่เหลือตามข้อ ๖.๒ แล้ว เว้นแต่ในกรณีดังต่อไปนี้ ผู้ขายมีสิทธิขอคืนหลักประกันการรับเงินล่วงหน้าบางส่วนก่อนได้<br/>
    (๑) กรณีผู้ขายได้วางหลักประกันการรับเงินล่วงหน้าไว้ฉบับเดียว หากผู้ซื้อได้หักเงินล่วงหน้าไปแล้ว ผู้ขายมีสิทธิขอคืนหลักประกันการรับเงินล่วงหน้าในส่วนที่ผู้ซื้อได้หักเงินล่วงหน้าไปแล้วนั้น โดยผู้ขายจะต้องนำหลักประกันการรับเงินล่วงหน้าฉบับใหม่ที่มีมูลค่าเท่ากับเงินล่วงหน้าที่เหลืออยู่มาวางให้แก่ผู้ซื้อ<br/>
    (๒) กรณีผู้ขายได้วางหลักประกันการรับเงินล่วงหน้าไว้หลายฉบับ ซึ่งแต่ละฉบับมีมูลค่าเท่ากับจำนวนเงินล่วงหน้าที่ผู้ซื้อจะต้องหักไว้ในแต่ละงวด หากผู้ซื้อได้หักเงินล่วงหน้าในงวดใดแล้วผู้ขายมีสิทธิขอคืนหลักประกันการรับเงินล่วงหน้าในงวดนั้นได้<br/>
    ๖.๒ เงินที่เหลือ จำนวน {result.RemainingPaymentAmount} บาท ({CommonDAO.NumberToThaiText(result.RemainingPaymentAmount ?? 0)}) จะจ่ายให้เมื่อผู้ซื้อได้รับคอมพิวเตอร์ตามข้อ ๕ ไว้โดยถูกต้องครบถ้วน และผู้ขายได้อบรมเจ้าหน้าที่ของผู้ซื้อตามข้อ ๑๐ เสร็จสิ้นแล้ว<br/>
    (11) การจ่ายเงินตามเงื่อนไขแห่งสัญญานี้ ผู้ซื้อจะโอนเงินเข้าบัญชีเงินฝากธนาคารของผู้ขาย ชื่อธนาคาร {result.SaleBankName} สาขา {result.SaleBankBranch} ชื่อบัญชี {result.SaleBankAccountName} เลขที่บัญชี {result.SaleBankAccountNumber} ทั้งนี้ ผู้ขายตกลงเป็นผู้รับภาระเงินค่าธรรมเนียม หรือค่าบริการอื่นใดเกี่ยวกับการโอน รวมทั้งค่าใช้จ่ายอื่นใด (ถ้ามี) ที่ธนาคารเรียกเก็บและยินยอมให้มีการหักเงินดังกล่าวจากจำนวนเงินโอนในงวดนั้น ๆ
</div>

<!-- ข้อ ๗ การรับประกันความชำรุดบกพร่อง -->
<div class='tab2 t-16'>ข้อ ๗ การรับประกันความชำรุดบกพร่อง</div>
<div class='tab3 t-16'>
    ผู้ขายตกลงรับประกันความชำรุดบกพร่องหรือขัดข้องของคอมพิวเตอร์และการติดตั้งตามสัญญานี้เป็นเวลา {result.WarrantyPeriodYears} ปี {result.WarrantyPeriodMonths} เดือน นับถัดจากวันที่ผู้ซื้อได้รับมอบคอมพิวเตอร์ทั้งหมดโดยถูกต้องครบถ้วนตามสัญญา ถ้าภายในระยะเวลาดังกล่าวคอมพิวเตอร์ชำรุดบกพร่องหรือขัดข้อง หรือใช้งานไม่ได้ทั้งหมดหรือแต่บางส่วน หรือเกิดความชำรุดบกพร่องหรือขัดข้องจากการติดตั้ง เว้นแต่ความชำรุดบกพร่องหรือขัดข้องดังกล่าว เกิดขึ้นจากความผิดของผู้ซื้อซึ่งไม่ได้เกิดขึ้นจากการใช้งานตามปกติ ผู้ขายจะต้องจัดการซ่อมแซมแก้ไขให้อยู่ในสภาพใช้การได้ดีดังเดิม โดยต้องเริ่มจัดการซ่อมแซมแก้ไขภายใน {result.DaysToRepairAfterNoti} วัน นับถัดจากวันที่ได้รับแจ้งจากผู้ซื้อโดยไม่คิดค่าใช้จ่ายใด ๆ จากผู้ซื้อทั้งสิ้น ถ้าผู้ขายไม่จัดการซ่อมแซมแก้ไขภายในกำหนดเวลาดังกล่าว ผู้ซื้อมีสิทธิที่จะทำการนั้นเองหรือจ้างผู้อื่นทำการนั้นแทนผู้ขาย โดยผู้ขายต้องออกค่าใช้จ่ายเองทั้งสิ้นแทนผู้ซื้อ<br/>
    ผู้ขายมีหน้าที่บำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ให้อยู่ในสภาพใช้งานได้ดีอยู่เสมอตลอดระยะเวลาดังกล่าวในวรรคหนึ่งด้วยค่าใช้จ่ายของผู้ขาย โดยให้มีเวลาคอมพิวเตอร์ขัดข้องรวมตามเกณฑ์การคำนวณเวลาขัดข้องไม่เกินเดือนละ {result.MaximumDownTimeHours} ชั่วโมง หรือร้อยละ {result.MaximumDownTimePercent}(%) ของเวลาใช้งานทั้งหมดของคอมพิวเตอร์ของเดือนนั้น แล้วแต่ตัวเลขใดจะมากกว่ากัน มิฉะนั้นผู้ขายต้องยอมให้ผู้ซื้อคิดค่าปรับเป็นรายชั่วโมง ในอัตราร้อยละ {result.PenaltyPerHourPercent} ของราคาคอมพิวเตอร์ทั้งหมดตามสัญญานี้คิดเป็นเงิน {result.PenaltyPerHour} บาท ({CommonDAO.NumberToThaiText(result.PenaltyPerHour ?? 0)}) ต่อชั่วโมง ในช่วงเวลาที่ไม่สามารถใช้คอมพิวเตอร์ได้ในส่วนที่เกินกว่ากำหนดเวลาขัดข้องข้างต้น<br/>
    เกณฑ์การคำนวณเวลาขัดข้องของคอมพิวเตอร์ตามวรรคสอง ให้เป็นดังนี้<br/>
    - กรณีที่คอมพิวเตอร์เกิดขัดข้องพร้อมกันหลายหน่วย ให้นับเวลาขัดข้องของหน่วยที่มีตัวถ่วงมากที่สุดเพียงหน่วยเดียว<br/>
    - กรณีความเสียหายอันสืบเนื่องมาจากความขัดข้องของคอมพิวเตอร์แตกต่างกัน เวลาที่ใช้ในการคำนวณค่าปรับจะเท่ากับเวลาขัดข้องของคอมพิวเตอร์หน่วยนั้นคูณด้วยตัวถ่วงซึ่งมีค่าต่างๆ ตามเอกสารแนบท้ายสัญญาผนวก ๔<br/>
    การที่ผู้ซื้อทำการนั้นเอง หรือให้ผู้อื่นทำการนั้นแทนผู้ขาย ไม่ทำให้ผู้ขายหลุดพ้นจาก ความรับผิดตามสัญญา หากผู้ขายไม่ชดใช้ค่าใช้จ่ายหรือค่าเสียหายตามที่ผู้ซื้อเรียกร้องผู้ซื้อมีสิทธิบังคับจากหลักประกันการปฏิบัติตามสัญญาได้
</div>

<!-- ข้อ ๘ หลักประกันการปฏิบัติตามสัญญา -->
<div class='tab2 t-16'>ข้อ ๘ หลักประกันการปฏิบัติตามสัญญา</div>
<div class='tab3 t-16'>
    ในขณะทำสัญญานี้ผู้ขายได้นำหลักประกันเป็น {result.PerformanceGuarantee} เป็นจำนวนเงิน {result.GuaranteeAmount} บาท ({CommonDAO.NumberToThaiText(result.GuaranteeAmount ?? 0)}) ซึ่งเท่ากับร้อยละ {result.GuaranteePercent} ของราคาซื้อขายคอมพิวเตอร์ตามข้อ ๑ มามอบให้แก่ผู้ซื้อเพื่อเป็นหลักประกันการปฏิบัติตามสัญญานี้<br/>
    (16) กรณีผู้ขายใช้หนังสือค้ำประกันมาเป็นหลักประกันการปฏิบัติตามสัญญา หนังสือค้ำประกันดังกล่าวจะต้องออกโดยธนาคารที่ประกอบกิจการในประเทศไทย หรือโดยบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบตามแบบที่คณะกรรมการนโยบายการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐกำหนด หรืออาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้และจะต้องมีอายุการค้ำประกันตลอดไปจนกว่าผู้ขายพ้นข้อผูกพันตามสัญญานี้<br/>
    หลักประกันที่ผู้ขายนำมามอบให้ตามวรรคหนึ่ง จะต้องมีอายุครอบคลุมความรับผิดทั้งปวงของผู้ขายตลอดอายุสัญญา ถ้าหลักประกันที่ผู้ขายนำมามอบให้ดังกล่าวลดลงหรือเสื่อมค่าลง หรือมีอายุไม่ครอบคลุมถึงความรับผิดของผู้ขายตลอดอายุสัญญา ไม่ว่าด้วยเหตุใดๆ ก็ตาม รวมถึงกรณีผู้ขายส่งมอบและติดตั้งคอมพิวเตอร์ล่าช้าเป็นเหตุให้ระยะเวลาแล้วเสร็จหรือวันครบกำหนดความรับผิดในความชำรุดบกพร่องตามสัญญาเปลี่ยนแปลงไป ไม่ว่าจะเกิดขึ้นคราวใด ผู้ขายต้องหาหลักประกันใหม่หรือหลักประกันเพิ่มเติมให้มีจำนวนครบถ้วนตามวรรคหนึ่ง มามอบให้แก่ผู้ซื้อภายใน {result.NewGuaranteeDays} วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ซื้อ<br/>
    หลักประกันที่ผู้ขายนำมามอบไว้ตามข้อนี้ ผู้ซื้อจะคืนให้แก่ผู้ขายโดยไม่มีดอกเบี้ยเมื่อผู้ขายพ้นจากข้อผูกพันและความรับผิดทั้งปวงตามสัญญานี้แล้ว
</div>

<!-- ข้อ ๙ การโอนกรรมสิทธิ์ -->
<div class='tab2 t-16'>ข้อ ๙ การโอนกรรมสิทธิ์</div>
<div class='tab3 t-16'>
    คู่สัญญาตกลงกันว่ากรรมสิทธิ์ในคอมพิวเตอร์ตามสัญญาจะโอนไปยังผู้ซื้อเมื่อผู้ซื้อได้รับมอบคอมพิวเตอร์ดังกล่าวตามข้อ ๕ แล้ว
</div>

<!-- ข้อ ๑๐ การอบรม -->
<div class='tab2 t-16'>ข้อ ๑๐ การอบรม</div>
<div class='tab3 t-16'>
    ผู้ขายต้องจัดอบรมวิชาการด้านคอมพิวเตอร์ให้แก่เจ้าหน้าที่ของผู้ซื้อจนสามารถใช้งานคอมพิวเตอร์ได้อย่างมีประสิทธิภาพ โดยต้องดำเนินการฝึกอบรมให้แล้วเสร็จภายใน {result.TrainingPeriodDays} วันนับถัดจากวันที่ผู้ซื้อได้รับมอบคอมพิวเตอร์ โดยไม่คิดค่าใช้จ่ายใด ๆ รายละเอียดของการฝึกอบรมให้เป็นไปตามเอกสารแนบท้ายสัญญาผนวก ๕
</div>

<!-- ข้อ ๑๑ คู่มือการใช้คอมพิวเตอร์ -->
<div class='tab2 t-16'>ข้อ ๑๑ คู่มือการใช้คอมพิวเตอร์</div>
<div class='tab3 t-16'>
    ผู้ขายต้องจัดหาและส่งมอบคู่มือการใช้คอมพิวเตอร์ตามสัญญานี้ ตามที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก ๖ จำนวน {result.ComputerManualsCount} ชุด ให้กับผู้ซื้อในวันที่ส่งมอบคอมพิวเตอร์ พร้อมทั้งปรับปรุงให้ทันสมัยเป็นปัจจุบันตลอดอายุสัญญานี้ ทั้งนี้ โดยไม่คิดเงินเพิ่มจากผู้ซื้อ
</div>

<!-- ข้อ ๑๒ การรับประกันความเสียหาย -->
<div class='tab2 t-16'>ข้อ ๑๒ การรับประกันความเสียหาย</div>
<div class='tab3 t-16'>
    ในกรณีที่บุคคลภายนอกกล่าวอ้างหรือใช้สิทธิเรียกร้องใด ๆ ว่ามีการละเมิดลิขสิทธิ์หรือสิทธิบัตรหรือสิทธิใด ๆ เกี่ยวกับคอมพิวเตอร์ตามสัญญานี้ โดยผู้ซื้อมิได้แก้ไขดัดแปลงไปจากเดิม ผู้ขายจะต้องดำเนินการทั้งปวงเพื่อให้การกล่าวอ้างหรือการเรียกร้องดังกล่าวระงับสิ้นไปโดยเร็ว หากผู้ซื้อต้องรับผิดชดใช้ค่าเสียหายต่อบุคคลภายนอกเนื่องจากผลแห่งการละเมิดลิขสิทธิ์หรือสิทธิบัตรหรือสิทธิใด ๆ ดังกล่าว ผู้ขายต้องเป็นผู้ชำระค่าเสียหายและค่าใช้จ่ายรวมทั้งค่าฤชาธรรมเนียมและค่าทนายความแทนผู้ซื้อ ทั้งนี้ ผู้ซื้อจะแจ้งให้ผู้ขายทราบเป็นหนังสือในเมื่อได้มีการกล่าวอ้างหรือใช้สิทธิเรียกร้องดังกล่าวโดยไม่ชักช้า
</div>

<!-- ข้อ ๑๓ การบอกเลิกสัญญา -->
<div class='tab2 t-16'>ข้อ ๑๓ การบอกเลิกสัญญา</div>
<div class='tab3 t-16'>
    เมื่อครบกำหนดส่งมอบคอมพิวเตอร์ตามสัญญาแล้ว หากผู้ขายไม่ส่งมอบและติดตั้งคอมพิวเตอร์บางรายการหรือทั้งหมดให้แก่ผู้ซื้อภายในกำหนดเวลาดังกล่าว หรือส่งมอบคอมพิวเตอร์ไม่ตรงตามสัญญาหรือมีคุณสมบัติไม่ถูกต้องตามสัญญา หรือส่งมอบและติดตั้งแล้วเสร็จภายในกำหนดแต่ไม่สามารถใช้งานได้อย่างมีประสิทธิภาพ หรือใช้งานไม่ได้ครบถ้วนตามสัญญา หรือผู้ขายไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่ง ผู้ซื้อมีสิทธิบอกเลิกสัญญาทั้งหมดหรือแต่บางส่วนได้ การใช้สิทธิบอกเลิกสัญญานั้นไม่กระทบถึงสิทธิของผู้ซื้อที่จะเรียกร้องค่าเสียหายจากผู้ขาย<br/>
    ในกรณีที่ผู้ซื้อใช้สิทธิบอกเลิกสัญญา ผู้ซื้อมีสิทธิริบหรือบังคับจากหลักประกันตาม (17) (ข้อ ๖ และ) ข้อ ๘ เป็นจำนวนเงินทั้งหมดหรือแต่บางส่วนก็ได้ แล้วแต่ผู้ซื้อจะเห็นสมควร และถ้าผู้ซื้อจัดซื้อคอมพิวเตอร์รวมถึงการติดตั้งจากบุคคลอื่นเต็มจำนวนหรือเฉพาะจำนวนที่ขาดส่ง แล้วแต่กรณี ภายในกำหนด {result.TeminationNewMonths} เดือน นับถัดจากวันบอกเลิกสัญญา ผู้ขายจะต้องชดใช้ราคาที่เพิ่มขึ้นจากราคาที่กำหนดไว้ในสัญญานี้ด้วย รวมทั้งค่าใช้จ่ายใด ๆ ที่ผู้ซื้อต้องใช้จ่ายในการจัดหาผู้ขายรายใหม่ดังกล่าวด้วย ในกรณีที่ผู้ขายได้ส่งมอบคอมพิวเตอร์ให้แก่ผู้ซื้อและผู้ซื้อบอกเลิกสัญญา ผู้ขายจะต้องนำคอมพิวเตอร์กลับคืนไป และทำสถานที่ที่รื้อถอนคอมพิวเตอร์ออกไปให้มีสภาพดังที่มีอยู่เดิมก่อนทำสัญญานี้ภายใน {result.ReturnDaysIn} วัน นับถัดจากวันที่ผู้ซื้อบอกเลิกสัญญา โดยผู้ขายเป็นผู้เสียค่าใช้จ่ายเองทั้งสิ้น<br/>
    (19)ถ้าผู้ขายไม่ยอมนำคอมพิวเตอร์กลับคืนไปภายในกำหนดเวลาดังกล่าวตามวรรคสาม ผู้ซื้อจะกำหนดเวลาให้ผู้ขายนำคอมพิวเตอร์กลับคืนไปอีกครั้งหนึ่ง หากพ้นกำหนดเวลาดังกล่าวแล้ว ผู้ขายยังไม่นำคอมพิวเตอร์กลับคืนไปอีก ผู้ซื้อมีสิทธินำคอมพิวเตอร์ออกขายทอดตลาด เงินที่ได้จากการขายทอดตลาด ผู้ขายยอมให้ผู้ซื้อหักเป็นค่าปรับและหักเป็นค่าใช้จ่าย และค่าเสียหายที่เกิดแก่ผู้ซื้อ ซึ่งรวมถึงค่าใช้จ่ายต่างๆ ที่ผู้ซื้อได้เสียไปในการดำเนินการขายทอดตลาดคอมพิวเตอร์ดังกล่าว ค่าใช้จ่ายในการทำสถานที่ที่รื้อถอนคอมพิวเตอร์ออกไปให้มีสภาพดังที่มีอยู่เดิมก่อนทำสัญญานี้ เงินที่เหลือจากการหักค่าปรับ ค่าใช้จ่าย และค่าเสียหายดังกล่าวแล้วผู้ซื้อจะคืนให้แก่ผู้ขาย<br/>
    เมื่อผู้ซื้อบอกเลิกสัญญาแล้ว ผู้ซื้อไม่ต้องรับผิดชอบในความเสียหายใด ๆ ทั้งสิ้นอันเกิดแก่คอมพิวเตอร์ซึ่งอยู่ในความครอบครองของผู้ซื้อ
</div>

<!-- ข้อ ๑๔ ค่าปรับ -->
<div class='tab2 t-16'>ข้อ ๑๔ ค่าปรับ</div>
<div class='tab3 t-16'>
    ในกรณีที่ผู้ซื้อยังไม่ใช้สิทธิบอกเลิกสัญญาตามข้อ ๑๓ ผู้ขายจะต้องชำระค่าปรับให้ผู้ซื้อเป็นรายวัน ในอัตราร้อยละ {result.FinePerDaysPercent} ของราคาคอมพิวเตอร์ที่ยังไม่ได้รับมอบ นับถัดจากวันครบกำหนดตามสัญญาจนถึงวันที่ผู้ขายได้นำคอมพิวเตอร์มาส่งมอบและติดตั้งให้แก่ผู้ซื้อจนถูกต้องครบถ้วนตามสัญญา<br/>
    การคิดค่าปรับในกรณีที่คอมพิวเตอร์ที่ตกลงซื้อขายเป็นระบบ ถ้าผู้ขายส่งมอบเพียงบางส่วนหรือขาดส่วนประกอบส่วนหนึ่งส่วนใดไป หรือส่งมอบและติดตั้งทั้งหมดแล้วแต่ใช้งานไม่ได้ถูกต้องครบถ้วน ให้ถือว่ายังไม่ได้ส่งมอบคอมพิวเตอร์นั้นเลย และคิดค่าปรับจากราคาคอมพิวเตอร์ทั้งระบบ<br/>
    ในระหว่างที่ผู้ซื้อยังไม่ได้ใช้สิทธิบอกเลิกสัญญานั้น หากผู้ซื้อเห็นว่าผู้ขายไม่อาจปฏิบัติตามสัญญาต่อไปได้ ผู้ซื้อจะใช้สิทธิบอกเลิกสัญญา และริบหรือบังคับจากหลักประกันตาม (21) (ข้อ ๖ และ) ข้อ ๘ กับเรียกร้องให้ชดใช้ราคาที่เพิ่มขึ้นตามที่กำหนดไว้ในข้อ ๑๓ วรรคสอง ก็ได้ และถ้าผู้ซื้อได้แจ้งข้อเรียกร้องให้ชำระค่าปรับไปยังผู้ขายเมื่อครบกำหนดส่งมอบแล้ว ผู้ซื้อมีสิทธิที่จะปรับผู้ขายจนถึงวันบอกเลิกสัญญาได้อีกด้วย
</div>

<!-- ข้อ ๑๕ การบังคับค่าปรับ ค่าเสียหาย และค่าใช้จ่าย -->
<div class='tab2 t-16'>ข้อ ๑๕ การบังคับค่าปรับ ค่าเสียหาย และค่าใช้จ่าย</div>
<div class='tab3 t-16'>
    ในกรณีที่ผู้ขายไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่งด้วยเหตุใด ๆ ก็ตาม จนเป็นเหตุให้เกิดค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายแก่ผู้ซื้อ ผู้ขายต้องชดใช้ค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายดังกล่าวให้แก่ผู้ซื้อ โดยสิ้นเชิงภายในกำหนด {result.EnforcementOfFineDays} วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ซื้อ หากผู้ขายไม่ชดใช้ให้ถูกต้องครบถ้วนภายในระยะเวลาดังกล่าวให้ผู้ซื้อมีสิทธิที่จะหักเอาจากค่าคอมพิวเตอร์ที่ต้องชำระ หรือบังคับจากหลักประกันการปฏิบัติตามสัญญาได้ทันที<br/>
    หากค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายที่บังคับจากค่าคอมพิวเตอร์ที่ต้องชำระหรือหลักประกันการปฏิบัติตามสัญญาแล้วยังไม่เพียงพอ ผู้ขายยินยอมชำระส่วนที่เหลือที่ยังขาดอยู่จนครบถ้วนตามจำนวนค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายนั้น ภายในกำหนด {result.OutstandingPeriodDays} วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ซื้อ<br/>
    หากมีเงินค่าคอมพิวเตอร์ที่ซื้อขายตามสัญญาที่หักไว้จ่ายเป็นค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายแล้วยังเหลืออยู่อีกเท่าใด ผู้ซื้อจะคืนให้แก่ผู้ขายทั้งหมด
</div>

<!-- ข้อ ๑๖ การงดหรือลดค่าปรับ หรือขยายเวลาในการปฏิบัติตามสัญญา -->
<div class='tab2 t-16'>ข้อ ๑๖ การงดหรือลดค่าปรับ หรือขยายเวลาในการปฏิบัติตามสัญญา</div>
<div class='tab3 t-16'>
    ในกรณีที่มีเหตุเกิดจากความผิดหรือความบกพร่องของผู้ซื้อ หรือเหตุสุดวิสัย หรือเกิดจากพฤติการณ์อันหนึ่งอันใดที่ผู้ขายไม่ต้องรับผิดตามกฎหมายหรือเหตุอื่นตามที่กำหนดในกฎกระทรวง ซึ่งออกตามความในกฎหมายว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ ทำให้ผู้ขายไม่สามารถส่งมอบและติดตั้งคอมพิวเตอร์ หรือบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ ตามเงื่อนไขและกำหนดเวลาแห่งสัญญานี้ได้ ผู้ขายมีสิทธิของดหรือลดค่าปรับหรือขยายเวลาทำการตามสัญญา โดยจะต้องแจ้งเหตุหรือพฤติการณ์ดังกล่าวพร้อมหลักฐานเป็นหนังสือให้ผู้ซื้อทราบภายใน ๑๕ (สิบห้า) วัน นับถัดจากวันที่เหตุนั้นสิ้นสุดลง หรือตามที่กำหนดในกฎกระทรวงดังกล่าว<br/>
    ถ้าผู้ขายไม่ปฏิบัติให้เป็นไปตามความในวรรคหนึ่ง ให้ถือว่าผู้ขายได้สละสิทธิเรียกร้องในการที่จะของดหรือลดค่าปรับหรือขยายเวลาทำการตามสัญญา โดยไม่มีเงื่อนไขใด ๆ ทั้งสิ้น เว้นแต่กรณีเหตุเกิดจากความผิดหรือความบกพร่องของฝ่ายผู้ซื้อซึ่งมีหลักฐานชัดแจ้งหรือผู้ซื้อทราบดีอยู่แล้วตั้งแต่ต้น<br/>
    การงดหรือลดค่าปรับหรือขยายเวลาทำการตามสัญญาตามวรรคหนึ่ง อยู่ในดุลพินิจของผู้ซื้อที่จะพิจารณาตามที่เห็นสมควร
</div>

<!-- ข้อ ๑๗ การใช้เรือไทย -->
<div class='tab2 t-16'>ข้อ ๑๗ การใช้เรือไทย</div>
<div class='tab3 t-16'>
    ถ้าผู้ขายจะต้องสั่งหรือนำเข้าคอมพิวเตอร์มาจากต่างประเทศและต้องนำเข้ามาโดยทางเรือในเส้นทางเดินเรือที่มีเรือไทยเดินอยู่ และสามารถให้บริการรับขนได้ตามที่รัฐมนตรีว่าการกระทรวงคมนาคมประกาศกำหนด ผู้ขายต้องจัดการให้คอมพิวเตอร์บรรทุกโดยเรือไทยหรือเรือที่มีสิทธิเช่นเดียวกับเรือไทยจากต่างประเทศมายังประเทศไทย เว้นแต่จะได้รับอนุญาตจากกรมเจ้าท่าก่อนบรรทุกคอมพิวเตอร์ลงเรืออื่นที่มิใช่เรือไทยหรือเป็นของที่รัฐมนตรีว่าการกระทรวงคมนาคมประกาศยกเว้นให้บรรทุกโดยเรืออื่นได้ ทั้งนี้ ไม่ว่าการสั่งหรือนำเข้าคอมพิวเตอร์จากต่างประเทศจะเป็นแบบใด<br/>
    ในการส่งมอบคอมพิวเตอร์ให้แก่ผู้ซื้อ ถ้าเป็นกรณีตามวรรคหนึ่งผู้ขายจะต้องส่งมอบใบตราส่ง (Bill of Lading) หรือสำเนาใบตราส่งสำหรับคอมพิวเตอร์นั้น ซึ่งแสดงว่าได้บรรทุกมาโดยเรือไทย หรือเรือที่มีสิทธิเช่นเดียวกับเรือไทยให้แก่ผู้ซื้อพร้อมกับการส่งมอบคอมพิวเตอร์ด้วย<br/>
    ในกรณีที่คอมพิวเตอร์ไม่ได้บรรทุกจากต่างประเทศมายังประเทศไทยโดยเรือไทยหรือเรือที่มีสิทธิเช่นเดียวกับเรือไทย ผู้ขายต้องส่งมอบหลักฐานซึ่งแสดงว่าได้รับอนุญาตจากกรมเจ้าท่าให้บรรทุกของโดยเรืออื่นได้ หรือหลักฐานซึ่งแสดงว่าได้ชำระค่าธรรมเนียมพิเศษเนื่องจากการไม่บรรทุกของโดยเรือไทยตามกฎหมายว่าด้วยการส่งเสริมการพาณิชยนาวีแล้วอย่างใดอย่างหนึ่งแก่ผู้ซื้อด้วย<br/>
    ในกรณีที่ผู้ขายไม่ส่งมอบหลักฐานอย่างใดอย่างหนึ่งดังกล่าวในวรรคสองและวรรคสามให้แก่ผู้ซื้อแต่จะขอส่งมอบคอมพิวเตอร์ให้ผู้ซื้อก่อนโดยยังไม่รับชำระเงินค่าคอมพิวเตอร์ ผู้ซื้อมีสิทธิรับคอมพิวเตอร์ไว้ก่อนและชำระเงินค่าคอมพิวเตอร์เมื่อผู้ขายได้ปฏิบัติถูกต้องครบถ้วนดังกล่าวแล้วได้<br/>
    สัญญานี้ทำขึ้นสองฉบับ มีข้อความถูกต้องตรงกัน คู่สัญญาได้อ่านและเข้าใจข้อความโดยละเอียดตลอดแล้ว จึงได้ลงลายมือชื่อพร้อมทั้งประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยานและคู่สัญญาต่างยึดถือไว้ฝ่ายละหนึ่งฉบับ
</div>

<!-- Signatures -->
<div class='contract text-center t-16'>
    ลงชื่อ {result.OSMEP_Signer ?? "xxxxxxxxxx"} ผู้ซื้อ<br/>
    ({result.OSMEP_Signer})<br/>
    ลงชื่อ {result.Contract_Signer} ผู้ขาย<br/>
    ({result.Contract_Signer})<br/>
    ลงชื่อ {result.OSMEP_Witness ?? "พยาน SME"} พยาน<br/>
    ({result.OSMEP_Witness ?? "พยาน SME"})<br/>
    ลงชื่อ {result.Contract_Witness ?? "พยานขาย"} พยาน<br/>
    ({result.Contract_Witness ?? "พยานขาย"})
</div>

<!-- Next page: Appendix instructions -->
<div style='page-break-before: always;'></div>
<div class='text-center t-22'><b>วิธีปฏิบัติเกี่ยวกับสัญญาซื้อขายคอมพิวเตอร์</b></div>
<div class=' t-16'>
    (1) ให้ระบุเลขที่สัญญาในปีงบประมาณหนึ่งๆ ตามลำดับ<br/>
    (2) ให้ระบุชื่อของหน่วยงานของรัฐที่เป็นนิติบุคคล เช่น กรม ก. หรือรัฐวิสาหกิจ ข. เป็นต้น<br/>
    (3) ให้ระบุชื่อและตำแหน่งของหัวหน้าหน่วยงานของรัฐที่เป็นนิติบุคคลนั้น หรือผู้ที่ได้รับมอบอำนาจ เช่น นาย ก. อธิบดีกรม………...… หรือ นาย ข. ผู้ได้รับมอบอำนาจจากอธิบดีกรม………......………..<br/>
    (4) ให้ระบุชื่อผู้รับจ้าง<br/>
    ก. กรณีนิติบุคคล เช่น ห้างหุ้นส่วนสามัญจดทะเบียน ห้างหุ้นส่วนจำกัด บริษัทจำกัด<br/>
    ข. กรณีบุคคลธรรมดา ให้ระบุชื่อและที่อยู่<br/>
    (5) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง<br/>
    (6) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง<br/>
    (๗) ให้ระบุยี่ห้อคอมพิวเตอร์ รุ่นคอมพิวเตอร์<br/>
    (๘) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง<br/>
    (9) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง<br/>
    (10) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง<br/>
    (11) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง<br/>
    (12) ระยะเวลารับประกันและระยะเวลาแก้ไขซ่อมแซมจะกำหนดเท่าใด แล้วแต่ลักษณะของสิ่งของที่ซื้อขายกัน โดยให้อยู่ในดุลพินิจของผู้ซื้อ ทั้งนี้ จะต้องประกาศให้ทราบในเอกสารประกวดราคาด้วย<br/>
    (13) ให้กำหนดในอัตราระหว่างร้อยละ ๐.๐๒๕ – ๐.๐๓๕ ของราคาตามสัญญาต่อชั่วโมง<br/>
    (14) “หลักประกัน” หมายถึง หลักประกันที่ผู้รับจ้างนำมามอบไว้แก่หน่วยงานของรัฐเมื่อลงนามในสัญญาเพื่อประกันความเสียหายที่อาจเกิดขึ้นจากการปฏิบัติตามสัญญา ดังนี้<br/>
    (๑)เงินสด<br/>
    (๒)เช็คหรือดราฟท์ที่ธนาคารเซ็นสั่งจ่าย ซึ่งเป็นเช็คหรือดราฟท์ลงวันที่ที่ใช้เช็คหรือดราฟท์นั้นชำระต่อเจ้าหน้าที่ หรือก่อนวันนั้นไม่เกิน ๓ วันทำการ<br/>
    (๓)หนังสือค้ำประกันของธนาคารภายในประเทศตามตัวอย่างที่คณะกรรมการนโยบายกำหนด โดยอาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้<br/>
    (๔)หนังสือค้ำประกันของบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบ โดยอนุโลมให้ใช้ตามตัวอย่างหนังสือค้ำประกันของธนาคารที่คณะกรรมการนโยบายกำหนด<br/>
    (๕)พันธบัตรรัฐบาลไทย<br/>
    (15) ให้กำหนดจำนวนเงินหลักประกันการปฏิบัติตามสัญญาตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ ข้อ ๑๖๘ ทั้งนี้ จะต้องประกาศให้ทราบในเอกสารประกวดราคาด้วย<br/>
    (16) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง<br/>
    (17) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง<br/>
    (18) กำหนดเวลาที่ผู้ซื้อจะซื้อสิ่งของจากแหล่งอื่นเมื่อบอกเลิกสัญญาและมีสิทธิเรียกเงินในส่วนที่เพิ่มขึ้นจากราคาที่กำหนดไว้ในสัญญานั้น ให้อยู่ในดุลพินิจของผู้ซื้อโดยตกลงกับผู้ขาย<br/>
    (19) ความในวรรคนี้อาจแก้ไขเปลี่ยนแปลงได้ตามความเหมาะสมถ้าหน่วยงานของรัฐผู้ทำสัญญาสามารถกำหนดมาตรการอื่นใดในสัญญา หรือกำหนดทางปฏิบัติเพื่อแก้ปัญหาที่ผู้ขายไม่ยอมนำคอมพิวเตอร์กลับคืนไปได้<br/>
    20) อัตราค่าปรับตามสัญญาข้อ ๑๔ ให้กำหนดเป็นรายวันในอัตราตายตัวระหว่างร้อยละ ๐.๐๑ – ๐.๒๐ ของราคาพัสดุที่ยังไม่ได้ส่งมอบ ตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ ข้อ ๑๖๒ ส่วนกรณีจะปรับร้อยละเท่าใด ให้อยู่ในดุลพินิจของหน่วยงานของรัฐผู้ซื้อที่จะพิจารณา โดยคำนึงถึงราคาและลักษณะของพัสดุที่ซื้อ ซึ่งอาจมีผลกระทบต่อการที่ผู้ขายจะหลีกเลี่ยงไม่ปฏิบัติตามสัญญา แต่ทั้งนี้ การที่จะกำหนดค่าปรับเป็นร้อยละเท่าใดจะต้องกำหนดไว้ในเอกสารเชิญชวนด้วย<br/>
    (21) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง
</div>
</br>
</br>
   
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
        .table {{ width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 28pt; }}
        .table th, .table td {{ border: 1px solid #000; padding: 8px; }}

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
{htmlContent}
</body>
</html>
";

        if (_pdfConverter == null)
            throw new Exception("PDF service is not available.");

        var doc = new HtmlToPdfDocument()
        {
            GlobalSettings = {
            PaperSize = PaperKind.A4,
            Orientation = DinkToPdf.Orientation.Portrait,
            Margins = new MarginSettings
            {
                Top = 20,
                Bottom = 20,
                Left = 20,
                Right = 20
            }
        },
            Objects = {
            new ObjectSettings() {
                HtmlContent = html,
                FooterSettings = new FooterSettings
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
    #endregion 4.1.1.2.11.สัญญาเช่าคอมพิวเตอร์ ร.309-60

}
