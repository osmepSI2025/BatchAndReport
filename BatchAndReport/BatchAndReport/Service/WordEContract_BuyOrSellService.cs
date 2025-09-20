using BatchAndReport.DAO;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Threading.Tasks;


public class WordEContract_BuyOrSellService
{
    private readonly WordServiceSetting _w;
    private readonly Econtract_Report_SPADAO _econtractReportSPADAO;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    private readonly EContractDAO _eContractDAO;
    private readonly E_ContractReportDAO _eContractReportDAO;
    public WordEContract_BuyOrSellService(WordServiceSetting ws
         , Econtract_Report_SPADAO econtractReportSPADAO
         , IConverter pdfConverter
        ,
EContractDAO eContractDAO
        ,
E_ContractReportDAO eContractReportDAO



        )
    {
        _w = ws;
        _econtractReportSPADAO = econtractReportSPADAO;
        _pdfConverter = pdfConverter; // กำหนดค่า PDF Converter
        _eContractDAO = eContractDAO;
        _eContractReportDAO = eContractReportDAO;
    }
    #region 4.1.1.2.11.สัญญาเช่าคอมพิวเตอร์ ร.309-60 SPA30560
    public async Task<byte[]> OnGetWordContact_BuyOrSellService(string id) 
    {

        var result = await _econtractReportSPADAO.GetSPAAsync(id);

        if (result == null)
        {
            throw new Exception("SPA data not found.");
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

                body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("แบบสัญญา", "000000", "36"));
                body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("สัญญาซื้อขาย", "000000", "36"));
                // 2. Document title and subtitle
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.RightParagraph("สัญญาเลขที่…"+result.SPAContractNumber+"."));


                // With this:
                // With this:
                string datestring = CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now);
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สัญญาฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.)\r\n"
                    + "ตำบล/แขวง ทุ่งสองห้อง อำเภอ/เขต หลักสี่\r\n" +
                "จังหวัด กรุงเทพ เมื่อ" + datestring + "\r\n" +
                "ระหว่าง " + result.Contract_Organization + "\r\n" +
                "โดย " + result.SignatoryName + "\r\n" +
                "ซึ่งต่อไปในสัญญานี้เรียกว่า “ผู้ซื้อ” ฝ่ายหนึ่ง กับ…" + result.ContractorName + "", null, "32"));

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
                    "สัญญานี้เรียกว่า “ผู้ให้เช่า” อีกฝ่ายหนึ่ง", null, "32"));

                }


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("คู่สัญญาได้ตกลงกันมีข้อความดังต่อไปนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑.ข้อตกลงซื้อขาย", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ซื้อตกลงซื้อและผู้ขายตกลงขาย "+result.ProductDescription+"" +
                    "จำนวน "+result.Quantity+"("+result.Unit+") เป็นราคาทั้งสิ้น "+result.TotalAmount+" บาท ("+CommonDAO.NumberToThaiText(result.TotalAmount??0) + ")ซึ่งได้รวมภาษีมูลค่าเพิ่มจำนวน "+result.VatAmount+" บาท ("+CommonDAO.NumberToThaiText(result.VatAmount ?? 0) + ") ตลอดจนภาษีอากรอื่นๆ และค่าใช้จ่ายทั้งปวงด้วยแล้ว", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๒การรับรองคุณภาพ", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ขายรับรองว่าสิ่งของที่ขายให้ตามสัญญานี้เป็นของแท้ ของใหม่ ไม่เคยใช้งานมาก่อน " +
                    "ไม่เป็นของเก่าเก็บ และมีคุณภาพและคุณสมบัติไม่ต่ำกว่าที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก………" +
                    "ในกรณีที่เป็นการซื้อสิ่งของซึ่งจะต้องมีการตรวจทดสอบ ผู้ขายรับรองว่า เมื่อตรวจทดสอบแล้วต้องมีคุณภาพและคุณสมบัติไม่ต่ำกว่าที่กำหนดไว้ตามสัญญานี้ด้วย", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๓เอกสารอันเป็นส่วนหนึ่งของสัญญา", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เอกสารแนบท้ายสัญญาดังต่อไปนี้ให้ถือเป็นส่วนหนึ่งของสัญญานี้", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๓.๑ ผนวก ๑ ……(รายการคุณลักษณะเฉพาะ)…….จำนวน…..…..(……….…) หน้า", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๓.๒ ผนวก ๒ …….…..(แค็ตตาล็อก) (๙)……….จำนวน…..…..(……….…) หน้า", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๓.๓ ผนวก ๓ ………...….(แบบรูป) (๑๐)………........จำนวน……....(………….) หน้า", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๓.๔ ผนวก ๔ ……………..(ใบเสนอราคา)…….…..…..จำนวน…...….(………….) หน้า", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("…………….……………….ฯลฯ….……..………….……", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ความใดในเอกสารแนบท้ายสัญญาที่ขัดหรือแย้งกับข้อความในสัญญานี้ ให้ใช้ข้อความในสัญญา" +
                    "นี้บังคับ และในกรณีที่เอกสารแนบท้ายสัญญาขัดแย้งกันเอง ผู้ขายจะต้องปฏิบัติตามคำวินิจฉัยของ ผู้ซื้อ คำวินิจฉัยของผู้ซื้อให้ถือเป็นที่สุด และผู้ขายไม่มีสิทธิเรียกร้องราคา ค่าเสียหาย หรือค่าใช้จ่ายใดๆเพิ่มเติมจากผู้ซื้อทั้งสิ้น", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ๔ การส่งมอบ", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ขายจะส่งมอบสิ่งของที่ซื้อขายตามสัญญาให้แก่ผู้ซื้อ ณ "+result.DeliveryLocation+" ภายใน"+ CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now) + " ให้ถูกต้องและครบถ้วนตามที่กำหนดไว้ในข้อ ๑ แห่งสัญญานี้ พร้อมทั้งหีบห่อหรือเครื่องรัดพันผูกโดยเรียบร้อย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("การส่งมอบสิ่งของตามสัญญานี้ ไม่ว่าจะเป็นการส่งมอบเพียงครั้งเดียว หรือส่งมอบหลายครั้ง ผู้ขายจะต้องแจ้งกำหนดเวลาส่งมอบแต่ละครั้งโดยทำเป็นหนังสือนำไปยื่นต่อผู้ซื้อ ณ "+result.DeliveryNotifyLocation+" ในวันและเวลาทำการของผู้ซื้อ ก่อนวันส่งมอบไม่น้อยกว่า "+result.DeliveryNotifyDays+" วันทำการของผู้ซื้อ", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ๕ การตรวจรับ", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เมื่อผู้ซื้อได้ตรวจรับสิ่งของที่ส่งมอบและเห็นว่าถูกต้องครบถ้วนตามสัญญาแล้ว ผู้ซื้อจะออกหลักฐานการรับมอบเป็นหนังสือไว้ให้ เพื่อผู้ขายนำมาเป็นหลักฐานประกอบการขอรับเงิน ค่าสิ่งของนั้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ถ้าผลของการตรวจรับปรากฏว่า สิ่งของที่ผู้ขายส่งมอบไม่ตรงตามข้อ ๑ ผู้ซื้อทรงไว้ซึ่งสิทธิที่จะไม่รับสิ่งของนั้น ในกรณีเช่นว่านี้ ผู้ขายต้องรีบนำสิ่งของนั้นกลับคืนโดยเร็วที่สุดเท่าที่จะทำได้และนำสิ่งของมาส่งมอบให้ใหม่ หรือต้องทำการแก้ไขให้ถูกต้องตามสัญญาด้วยค่าใช้จ่ายของผู้ขายเอง และระยะเวลาที่เสียไปเพราะเหตุดังกล่าวผู้ขายจะนำมาอ้างเป็นเหตุขอขยายเวลาส่งมอบตามสัญญาหรือ ของดหรือลดค่าปรับไม่ได้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๑๒) ในกรณีที่ผู้ขายส่งมอบสิ่งของถูกต้องแต่ไม่ครบจำนวน หรือส่งมอบครบจำนวน แต่ไม่ถูกต้องทั้งหมด ผู้ซื้อจะตรวจรับเฉพาะส่วนที่ถูกต้อง โดยออกหลักฐานการตรวจรับเฉพาะส่วนนั้นก็ได้ (ความในวรรคสามนี้ จะไม่กำหนดไว้ในกรณีที่ผู้ซื้อต้องการสิ่งของทั้งหมดในคราวเดียวกัน หรือการซื้อสิ่งของที่ประกอบเป็นชุดหรือหน่วย ถ้าขาดส่วนประกอบอย่างหนึ่งอย่างใดไปแล้ว จะไม่สามารถใช้งานได้ โดยสมบูรณ์)", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ๖ การชำระเงิน", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs(" (13 ก) ผู้ซื้อตกลงชำระเงินค่าสิ่งของตามข้อ ๑ ให้แก่ผู้ขาย เมื่อผู้ซื้อได้รับมอบสิ่งของตามข้อ ๕ ไว้โดยครบถ้วนแล้ว", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs(" (13 ข) ผู้ซื้อตกลงชำระเงินค่าสิ่งของตามข้อ ๑ ให้แก่ผู้ขาย ดังนี้ ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๖.๑ เงินล่วงหน้า จำนวน "+result.AdvancePayment+" บาท ("+CommonDAO.NumberToThaiText(result.AdvancePayment??0) + ") จะจ่ายให้ภายใน "+result.PaymentDueDays+"" +
                    "วัน นับถัดจากวันลงนามในสัญญา ทั้งนี้ โดยผู้ขายจะต้องนำหลักประกันเงินล่วงหน้าเป็น"+result.PaymentGuaranteeType+"(หนังสือค้ำประกันหรือหนังสือค้ำประกันอิเล็กทรอนิกส์ของธนาคารภายในประเทศหรือพันธบัตรรัฐบาลไทย)….....เต็มตามจำนวนเงินล่วงหน้าที่จะได้รับ" +
                    "มามอบให้แก่ผู้ซื้อเป็นหลักประกันการชำระคืนเงินล่วงหน้าก่อนการรับชำระเงินล่วงหน้านั้น และผู้ซื้อจะคืนหลักประกันเงินล่วงหน้าให้แก่ผู้ขายเมื่อผู้ซื้อจ่ายเงินที่เหลือตามข้อ ๖.๒ ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๖.๒ เงินที่เหลือ จำนวน " + result.AdvancePayment + " บาท  (" + CommonDAO.NumberToThaiText(result.AdvancePayment??0) + ") จะจ่ายให้เมื่อผู้ซื้อได้รับมอบสิ่งของ ตามข้อ ๕ ไว้โดยถูกต้องครบถ้วนแล้ว", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๑๔) การจ่ายเงินตามเงื่อนไขแห่งสัญญานี้ ผู้ซื้อจะโอนเงินเข้าบัญชีเงินฝากธนาคารของผู้ขาย ชื่อธนาคาร "+result.SaleBankName+" สาขา "+result.SaleBankBranch+" ชื่อบัญชี "+result.SaleBankAccountName+" เลขที่บัญชี"+result.SaleBankAccountNumber+" ทั้งนี้ ผู้ขายตกลงเป็นผู้รับภาระเงินค่าธรรมเนียม หรือค่าบริการอื่นใดเกี่ยวกับการโอน รวมทั้งค่าใช้จ่ายใดๆ (ถ้ามี) ที่ธนาคารเรียกเก็บ และยินยอมให้มีการหักเงินดังกล่าวจากจำนวนเงินโอนในงวดนั้นๆ (ความในวรรคนี้ใช้สำหรับกรณีที่หน่วยงานของรัฐจะจ่ายเงินตรงให้แก่ผู้ขาย (ระบบ Direct Payment) โดยการโอนเงินเข้าบัญชีเงินฝากธนาคารของผู้ขาย ตามแนวทางที่กระทรวงการคลังหรือหน่วยงานของรัฐเจ้าของงบประมาณเป็นผู้กำหนด แล้วแต่กรณี)", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๗การรับประกันความชำรุดบกพร่อง", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ขายตกลงรับประกันความชำรุดบกพร่องหรือขัดข้องของสิ่งของตามสัญญานี้ เป็นเวลา...........(๑๕)............(……..…………….) ปี .…....….......…..(…….……..…….….)เดือน " +
                    "นับถัดจากวันที่ผู้ซื้อได้รับมอบสิ่งของทั้งหมดไว้โดยถูกต้องครบถ้วนตามสัญญา โดยภายในกำหนดเวลาดังกล่าว หากสิ่งของ ตามสัญญานี้เกิดชำรุดบกพร่องหรือขัดข้องอันเนื่องมาจากการใช้งานตามปกติ " +
                    "ผู้ขายจะต้องจัดการซ่อมแซมหรือแก้ไขให้อยู่ในสภาพที่ใช้การได้ดีดังเดิม ภายใน…….......(……..….) วัน นับถัดจากวันที่ได้รับแจ้งจากผู้ซื้อ โดยไม่คิดค่าใช้จ่ายใดๆ ทั้งสิ้น หากผู้ขายไม่จัดการซ่อมแซมหรือแก้ไขภายในกำหนดเวลาดังกล่าว " +
                    "ผู้ซื้อมีสิทธิที่จะทำการนั้นเองหรือจ้างผู้อื่นให้ทำการนั้นแทนผู้ขาย โดยผู้ขายต้องเป็นผู้ออกค่าใช้จ่ายเองทั้งสิ้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีเร่งด่วนจำเป็นต้องรีบแก้ไขเหตุชำรุดบกพร่องหรือขัดข้องโดยเร็ว " +
                    "และไม่อาจรอคอยให้ผู้ขายแก้ไขในระยะเวลาที่กำหนดไว้ตามวรรคหนึ่งได้ ผู้ซื้อมีสิทธิเข้าจัดการแก้ไขเหตุชำรุดบกพร่องหรือขัดข้องนั้นเอง " +
                    "หรือให้ผู้อื่นแก้ไขความชำรุดบกพร่องหรือขัดข้อง โดยผู้ขายต้องรับผิดชอบชำระค่าใช้จ่ายทั้งหมด", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("การที่ผู้ซื้อทำการนั้นเอง หรือให้ผู้อื่นทำการนั้นแทนผู้ขาย ไม่ทำให้ผู้ขายหลุดพ้นจากความรับผิดตามสัญญา หากผู้ขายไม่ชดใช้ค่าใช้จ่ายหรือค่าเสียหายตามที่ผู้ซื้อเรียกร้องผู้ซื้อมีสิทธิบังคับจากหลักประกันการปฏิบัติตามสัญญาได้", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๘หลักประกันการปฏิบัติตามสัญญา", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในขณะทำสัญญานี้ผู้ขายได้นำหลักประกันเป็น "+result.GuaranteeType+" เป็นจำนวนเงิน "+result.GuaranteeAmount+" บาท (" + CommonDAO.NumberToThaiText(result.GuaranteeAmount ?? 0) + ") ซึ่งเท่ากับร้อยละ "+result.GuaranteePercent+"(%) ของราคาทั้งหมดตามสัญญา มามอบให้แก่ผู้ซื้อเพื่อเป็นหลักประกันการปฏิบัติตามสัญญานี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๑๘) กรณีผู้ขายใช้หนังสือค้ำประกันมาเป็นหลักประกันการปฏิบัติตามสัญญา หนังสือค้ำประกันดังกล่าวจะต้องออกโดยธนาคารที่ประกอบกิจการในประเทศไทย หรือโดยบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบตามแบบที่คณะกรรมการนโยบายการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐกำหนด หรืออาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้ และจะต้องมีอายุการค้ำประกันตลอดไปจนกว่าผู้ขายพ้นข้อผูกพันตามสัญญานี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หลักประกันที่ผู้ขายนำมามอบให้ตามวรรคหนึ่ง จะต้องมีอายุครอบคลุมความรับผิดทั้งปวงของผู้ขายตลอดอายุสัญญานี้ ถ้าหลักประกันที่ผู้ขายนำมามอบให้ดังกล่าวลดลงหรือเสื่อมค่าลง หรือมีอายุไม่ครอบคลุมถึงความรับผิดของผู้ขายตลอดอายุสัญญา ไม่ว่าด้วยเหตุใดๆ ก็ตาม รวมถึงกรณีผู้ขายส่งมอบสิ่งของล่าช้าเป็นเหตุให้ระยะเวลาส่งมอบหรือวันครบกำหนดความรับผิดในความชำรุดบกพร่องตามสัญญาเปลี่ยนแปลงไป ไม่ว่าจะเกิดขึ้นคราวใด ผู้ขายต้องหาหลักประกันใหม่หรือหลักประกันเพิ่มเติมให้มีจำนวนครบถ้วนตามวรรคหนึ่งมามอบให้แก่ผู้ซื้อภายใน "+result.NewGuaranteeDays+" วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ซื้อ ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หลักประกันที่ผู้ขายนำมามอบไว้ตามข้อนี้ ผู้ซื้อจะคืนให้แก่ผู้ขายโดยไม่มีดอกเบี้ยเมื่อผู้ขายพ้นจากข้อผูกพันและความรับผิดทั้งปวงตามสัญญานี้แล้ว", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๙การบอกเลิกสัญญา", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ถ้าผู้ขายไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่ง หรือเมื่อครบกำหนดส่งมอบสิ่งของตามสัญญานี้แล้ว หากผู้ขายไม่ส่งมอบสิ่งของที่ตกลงขายให้แก่ผู้ซื้อหรือส่งมอบไม่ถูกต้อง หรือไม่ครบจำนวน ผู้ซื้อมีสิทธิบอกเลิกสัญญาทั้งหมดหรือแต่บางส่วนได้ การใช้สิทธิบอกเลิกสัญญานั้นไม่กระทบสิทธิของผู้ซื้อที่จะเรียกร้องค่าเสียหายจากผู้ขาย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้ซื้อใช้สิทธิบอกเลิกสัญญา ผู้ซื้อมีสิทธิริบหรือบังคับจากหลักประกันตาม (๑๙) (ข้อ ๖ และ) ข้อ ๘ เป็นจำนวนเงินทั้งหมดหรือแต่บางส่วนก็ได้ แล้วแต่ผู้ซื้อจะเห็นสมควร และถ้าผู้ซื้อจัดซื้อสิ่งของจากบุคคลอื่นเต็มจำนวนหรือเฉพาะจำนวนที่ขาดส่ง แล้วแต่กรณี ภายในกำหนด "+result.TerminationNewMonths+" เดือน นับถัดจากวันบอกเลิกสัญญา ผู้ขายจะต้องชดใช้ราคาที่เพิ่มขึ้นจากราคาที่กำหนดไว้ในสัญญานี้ด้วย", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๐ค่าปรับ", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้ซื้อมิได้ใช้สิทธิบอกเลิกสัญญาตามข้อ ๙ ผู้ขายจะต้องชำระค่าปรับให้ผู้ซื้อเป็นรายวันในอัตราร้อยละ "+result.FineRatePerDay+"(%) ของราคาสิ่งของที่ยังไม่ได้รับมอบ นับถัดจากวันครบกำหนดตามสัญญาจนถึงวันที่ผู้ขายได้นำสิ่งของมาส่งมอบให้แก่ผู้ซื้อจนถูกต้องครบถ้วนตามสัญญา ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("การคิดค่าปรับในกรณีสิ่งของที่ตกลงซื้อขายประกอบกันเป็นชุด แต่ผู้ขายส่งมอบเพียงบางส่วน หรือขาดส่วนประกอบส่วนหนึ่งส่วนใดไปทำให้ไม่สามารถใช้การได้โดยสมบูรณ์ ให้ถือว่า ยังไม่ได้ส่งมอบสิ่งของนั้นเลย และให้คิดค่าปรับจากราคาสิ่งของเต็มทั้งชุด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในระหว่างที่ผู้ซื้อยังมิได้ใช้สิทธิบอกเลิกสัญญานั้น หากผู้ซื้อเห็นว่าผู้ขายไม่อาจปฏิบัติตามสัญญาต่อไปได้ ผู้ซื้อจะใช้สิทธิบอกเลิกสัญญาและริบหรือบังคับจากหลักประกันตาม (๒๒) (ข้อ ๖ และ) ข้อ ๘ กับเรียกร้องให้ชดใช้ราคาที่เพิ่มขึ้นตามที่กำหนดไว้ในข้อ ๙ วรรคสองก็ได้ และถ้าผู้ซื้อได้แจ้งข้อเรียกร้องให้ชำระค่าปรับไปยังผู้ขายเมื่อครบกำหนดส่งมอบแล้ว ผู้ซื้อมีสิทธิที่จะปรับผู้ขายจนถึงวันบอกเลิกสัญญาได้อีกด้วย", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๑การบังคับค่าปรับ ค่าเสียหาย และค่าใช้จ่าย", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้ขายไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่งด้วยเหตุใดๆ ก็ตาม จนเป็นเหตุให้เกิดค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายแก่ผู้ซื้อ ผู้ขายต้องชดใช้ค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายดังกล่าวให้แก่ผู้ซื้อโดยสิ้นเชิงภายในกำหนด................(.................) วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ซื้อ หากผู้ขายไม่ชดใช้ให้ถูกต้องครบถ้วนภายในระยะเวลาดังกล่าวให้ผู้ซื้อมีสิทธิที่จะหักเอาจากจำนวนเงินค่าสิ่งของที่ซื้อขายที่ต้องชำระ หรือบังคับจากหลักประกันการปฏิบัติตามสัญญาได้ทันที", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หากค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายที่บังคับจากเงินค่าสิ่งของที่ซื้อขายที่ต้องชำระ หรือหลักประกันการปฏิบัติตามสัญญาแล้วยังไม่เพียงพอ ผู้ขายยินยอมชำระส่วนที่เหลือที่ยังขาดอยู่จนครบถ้วนตามจำนวนค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายนั้น ภายในกำหนด.................(..................) วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ซื้อ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หากมีเงินค่าสิ่งของที่ซื้อขายตามสัญญาที่หักไว้จ่ายเป็นค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายแล้วยังเหลืออยู่อีกเท่าใด ผู้ซื้อจะคืนให้แก่ผู้ขายทั้งหมด", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๒การงดหรือลดค่าปรับ หรือขยายเวลาส่งมอบ", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่มีเหตุเกิดจากความผิดหรือความบกพร่องของฝ่ายผู้ซื้อ หรือเหตุสุดวิสัย หรือเกิดจากพฤติการณ์อันหนึ่งอันใดที่ผู้ขายไม่ต้องรับผิดตามกฎหมาย หรือเหตุอื่นตามที่กำหนดในกฎกระทรวง ซึ่งออกตามความในกฎหมายว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ ทำให้ผู้ขายไม่สามารถส่งมอบสิ่งของตามเงื่อนไขและกำหนดเวลาแห่งสัญญานี้ได้ ผู้ขายมีสิทธิของดหรือลดค่าปรับหรือขยายเวลาส่งมอบตามสัญญาได้ โดยจะต้องแจ้งเหตุหรือพฤติการณ์ดังกล่าวพร้อมหลักฐานเป็นหนังสือให้ผู้ซื้อทราบภายใน ๑๕ (สิบห้า) วัน นับถัดจากวันที่เหตุนั้นสิ้นสุดลง หรือตามที่กำหนดในกฎกระทรวงดังกล่าว", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ถ้าผู้ขายไม่ปฏิบัติให้เป็นไปตามความในวรรคหนึ่ง ให้ถือว่าผู้ขายได้สละสิทธิเรียกร้องในการที่จะของดหรือลดค่าปรับหรือขยายเวลาส่งมอบตามสัญญา โดยไม่มีเงื่อนไขใดๆ ทั้งสิ้น เว้นแต่กรณีเหตุเกิดจากความผิดหรือความบกพร่องของฝ่ายผู้ซื้อซึ่งมีหลักฐานชัดแจ้งหรือผู้ซื้อทราบดีอยู่แล้วตั้งแต่ต้นการงดหรือลดค่าปรับหรือขยายเวลาส่งมอบตามสัญญาตามวรรคหนึ่ง อยู่ในดุลพินิจของผู้ซื้อที่จะพิจารณาตามที่เห็นสมควร", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๓การใช้เรือไทย", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ถ้าสิ่งของที่จะต้องส่งมอบให้แก่ผู้ซื้อตามสัญญานี้ เป็นสิ่งของที่ผู้ขายจะต้องสั่งหรือนำเข้ามาจากต่างประเทศ และสิ่งของนั้นต้องนำเข้ามาโดยทางเรือในเส้นทางเดินเรือที่มีเรือไทยเดินอยู่ และสามารถให้บริการรับขนได้ตามที่รัฐมนตรีว่าการกระทรวงคมนาคมประกาศกำหนด ผู้ขายต้องจัดการให้สิ่งของดังกล่าวบรรทุกโดยเรือไทยหรือเรือที่มีสิทธิเช่นเดียวกับเรือไทยจากต่างประเทศมายังประเทศไทย เว้นแต่จะได้รับอนุญาตจากกรมเจ้าท่าก่อนบรรทุกของนั้นลงเรืออื่นที่มิใช่เรือไทยหรือเป็นของที่รัฐมนตรี ว่าการกระทรวงคมนาคมประกาศยกเว้นให้บรรทุกโดยเรืออื่นได้ ทั้งนี้ ไม่ว่าการสั่งหรือนำเข้าสิ่งของดังกล่าวจากต่างประเทศจะเป็นแบบใด ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในการส่งมอบสิ่งของตามสัญญาให้แก่ผู้ซื้อ ถ้าสิ่งของนั้นเป็นสิ่งของตามวรรคหนึ่ง ผู้ขายจะต้องส่งมอบใบตราส่ง (Bill of Lading) หรือสำเนาใบตราส่งสำหรับของนั้น ซึ่งแสดงว่าได้บรรทุกมาโดยเรือไทยหรือเรือที่มีสิทธิเช่นเดียวกับเรือไทยให้แก่ผู้ซื้อพร้อมกับการส่งมอบสิ่งของด้วย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่สิ่งของดังกล่าวไม่ได้บรรทุกจากต่างประเทศมายังประเทศไทย โดยเรือไทยหรือเรือที่มีสิทธิเช่นเดียวกับเรือไทย ผู้ขายต้องส่งมอบหลักฐานซึ่งแสดงว่าได้รับอนุญาตจากกรมเจ้าท่า ให้บรรทุกของโดยเรืออื่นได้หรือหลักฐานซึ่งแสดงว่าได้ชำระค่าธรรมเนียมพิเศษเนื่องจากการไม่บรรทุกของโดยเรือไทยตามกฎหมายว่าด้วยการส่งเสริมการพาณิชยนาวีแล้วอย่างใดอย่างหนึ่งแก่ผู้ซื้อด้วย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้ขายไม่ส่งมอบหลักฐานอย่างใดอย่างหนึ่งดังกล่าวในวรรคสองและวรรคสามให้แก่ผู้ซื้อ แต่จะขอส่งมอบสิ่งของดังกล่าวให้ผู้ซื้อก่อนโดยยังไม่รับชำระเงินค่าสิ่งของ ผู้ซื้อมีสิทธิรับสิ่งของดังกล่าวไว้ก่อนและชำระเงินค่าสิ่งของเมื่อผู้ขายได้ปฏิบัติถูกต้องครบถ้วนดังกล่าวแล้วได้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("สัญญานี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน คู่สัญญาได้อ่านและเข้าใจข้อความ โดยละเอียดตลอดแล้ว จึงได้ลงลายมือชื่อพร้อมทั้งประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยาน และคู่สัญญาต่างยึดถือไว้ฝ่ายละหนึ่งฉบับ", null, "32"));



                body.AppendChild(WordServiceSetting.EmptyParagraph());

                body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ......"+result.OSMEP_Signer+"......ผู้ซื้อ"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(................................................................................)"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ......"+result.Contract_Signer+"..........................................ผู้ขาย"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(................................................................................)"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ........"+result.OSMEP_Witness+".........พยาน"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(...............................................................................)"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ........" + result.Contract_Witness + ".........พยาน"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(...............................................................................)"));

                // next page
                body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("วิธีปฏิบัติเกี่ยวกับสัญญาซื้อขายคอมพิวเตอร์", "000000", "36"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๑)  ให้ระบุเลขที่สัญญาในปีงบประมาณหนึ่งๆ ตามลำดับ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๒)  ให้ระบุชื่อของหน่วยงานของรัฐที่เป็นนิติบุคคล เช่น กรม ก. หรือรัฐวิสาหกิจ ข. เป็นต้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๓)  ให้ระบุชื่อและตำแหน่งของหัวหน้าหน่วยงานของรัฐที่เป็นนิติบุคคลนั้น หรือผู้ที่ได้รับมอบอำนาจ เช่น นาย ก. อธิบดีกรม……………… หรือ นาย ข. ผู้ได้รับมอบอำนาจจากอธิบดีกรม………………..", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๔)  ให้ระบุชื่อผู้ขาย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ก. กรณีนิติบุคคล เช่น ห้างหุ้นส่วนสามัญจดทะเบียน ห้างหุ้นส่วนจำกัด บริษัทจำกัด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ข. กรณีบุคคลธรรมดา ให้ระบุชื่อและที่อยู่", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๕)  เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๖)  เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๗)  ให้ระบุว่าเป็นการซื้อสิ่งของตามตัวอย่าง หรือรายการละเอียด หรือแค็ตตาล็อก หรือแบบรูปรายการ หรืออื่นๆ (ให้ระบุ) และปกติจะต้องกำหนดไว้ด้วยว่าสิ่งของที่จะซื้อนั้น เป็นของแท้ เป็นของใหม่ ไม่เคยใช้งานมาก่อน ", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๘)  ให้ระบุหน่วยที่ใช้ เช่น กิโลกรัม ชิ้น เมตร เป็นต้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๙)  เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๑๐)  เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๑๑)  กำหนดเวลาส่งมอบจะต้องแจ้งล่วงหน้าไม่น้อยกว่ากี่วัน ให้อยู่ในดุลพินิจของผู้ซื้อโดยตกลงกับผู้ขาย โดยปกติควรจะกำหนดไว้ประมาณ ๓ วันทำการ เพื่อที่ผู้ซื้อจะได้จัดเตรียมเจ้าหน้าที่ไว้ตรวจรับของนั้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ในกรณีที่มีการส่งมอบสิ่งของหลายครั้ง ให้ระบุวันเวลาที่ส่งมอบแต่ละครั้งไว้ด้วย และในกรณีที่มีการติดตั้งด้วย ให้แยกกำหนดเวลาส่งมอบ และกำหนดเวลาการติดตั้งออกจากกัน", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๑๒)  เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๑๓)  ให้หน่วยงานของรัฐเลือกใช้ตามความเหมาะสม", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ข้อความในข้อ 6 กรณีไม่มีการจ่ายเงินล่วงหน้าให้ผู้ขาย ให้เลือกใช้ข้อความในข้อ (13 ก)", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ข้อความในข้อ 6 กรณีมีการจ่ายเงินล่วงหน้าให้ผู้ขาย  ให้เลือกใช้ข้อความในข้อ (13 ข)", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๑๔)  เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๑๕)  ระยะเวลารับประกันและระยะเวลาแก้ไขซ่อมแซมจะกำหนดเท่าใด แล้วแต่ลักษณะของสิ่งของที่ซื้อขายกัน โดยให้อยู่ในดุลพินิจของผู้ซื้อ เช่น เครื่องคำนวณไฟฟ้า กำหนดเวลารับประกัน ๑ ปี กำหนดเวลาแก้ไขภายใน ๗ วัน เป็นต้น ทั้งนี้ จะต้องประกาศให้ทราบในเอกสารเชิญชวนด้วย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๑๖)  “หลักประกัน” หมายถึง หลักประกันที่ผู้ขายนำมามอบไว้แก่หน่วยงานของรัฐ เมื่อลงนามในสัญญา เพื่อเป็นการประกันความเสียหายที่อาจจะเกิดขึ้นจากการปฏิบัติตามสัญญา ดังนี้", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๑)เงินสด ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๒)เช็คหรือดราฟท์ที่ธนาคารเซ็นสั่งจ่าย ซึ่งเป็นเช็คหรือดราฟท์ลงวันที่ที่ใช้เช็คหรือดราฟท์นั้นชำระต่อเจ้าหน้าที่ หรือก่อนวันนั้นไม่เกิน ๓ วันทำการ ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๓)หนังสือค้ำประกันของธนาคารภายในประเทศตามตัวอย่างที่คณะกรรมการนโยบายกำหนด โดยอาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๔)หนังสือค้ำประกันของบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบ โดยอนุโลมให้ใช้ตามตัวอย่างหนังสือ ค้ำประกันของธนาคารที่คณะกรรมการนโยบายกำหนด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๕)พันธบัตรรัฐบาลไทย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๑๗)  ให้กำหนดจำนวนเงินหลักประกันการปฏิบัติตามสัญญาตามระเบียบกระทรวงการคลังว่าด้วยหลักเกณฑ์การจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ ข้อ ๑๖๘", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๑๘)  เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๑๙)  เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๒๐)  กำหนดเวลาที่ผู้ซื้อจะซื้อสิ่งของจากแหล่งอื่นเมื่อบอกเลิกสัญญาและมีสิทธิเรียกเงินในส่วนที่เพิ่มขึ้นจากราคาที่กำหนดไว้ในสัญญานั้น ให้อยู่ในดุลพินิจของผู้ซื้อโดยตกลงกับผู้ขาย และโดยปกติแล้วไม่ควรเกิน ๓ เดือน", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๒๑)  อัตราค่าปรับตามสัญญาข้อ 10 ให้กำหนดเป็นรายวันในอัตราระหว่างร้อยละ ๐.๑๐-๐.๒๐ ตามระเบียบกระทรวงการคลังว่าด้วยหลักเกณฑ์การจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ข้อ ๑๖๒ ส่วนกรณีจะปรับร้อยละเท่าใด ให้อยู่ในดุลพินิจของหน่วยงานของรัฐผู้ซื้อที่จะพิจารณา โดยคำนึงถึงราคาและลักษณะของพัสดุที่ซื้อ ซึ่งอาจมีผลกระทบต่อการที่ผู้ขายจะหลีกเลี่ยงไม่ปฏิบัติตามสัญญา แต่ทั้งนี้การที่จะกำหนดค่าปรับเป็นร้อยละเท่าใด จะต้องกำหนดไว้ในเอกสารเชิญชวนด้วย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(๒๒)  เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));

                body.AppendChild(WordServiceSetting.EmptyParagraph());




                WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);

            }
            stream.Position = 0;
            return stream.ToArray();

        }
 
    }
    public async Task<string> OnGetWordContact_BuyOrSellService_ToPDF(string id,string typeContact)
    {
        var result = await _econtractReportSPADAO.GetSPAAsync(id);
      
        if (result == null)
        {
            throw new Exception("SPA data not found.");
        }
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf");
        string fontBase64 = "";
        if (File.Exists(fontPath))
        {
            var bytes = File.ReadAllBytes(fontPath);
            fontBase64 = Convert.ToBase64String(bytes);
        }
        var cssPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "css", "contract.css").Replace("\\", "/");
        var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
        string logoBase64 = "";
        if (System.IO.File.Exists(logoPath))
        {
            var bytes = System.IO.File.ReadAllBytes(logoPath);
            logoBase64 = Convert.ToBase64String(bytes);
        }
        #region signlist 

        var signlist = await _eContractReportDAO.GetSignNameAsync(id, typeContact);
        var signatoryTableHtml = "";
        if (signlist.Count > 0)
        {
            signatoryTableHtml = await _eContractReportDAO.RenderSignatory(signlist);

        }
        #endregion signlist



        var strDateTH = CommonDAO.ToThaiDateString(result.ContractSignDate ?? DateTime.Now);
        //เอกสารแนบท้ายสัญญาผนวก 1-6
        var listDocAtt = await _eContractDAO.GetRelatedDocumentsAsync(id, "SPA30560");
        var htmlDocAtt = "";
        if (listDocAtt != null)
        {
            for (int i = 0; i < listDocAtt.Count; i++)
            {
                var docItem = listDocAtt[i];
                htmlDocAtt += $"<p class='tab3 t-14'> {docItem.DocumentTitle} จำนวน {docItem.PageAmount} หน้า </div>";


            }
        }

        var html = $@"<html>
<head>
    <meta charset='utf-8'>
  
        <style>
        @font-face {{
            font-family: 'THSarabunNew';
              src: url('data:font/truetype;charset=utf-8;base64,{fontBase64}') format('truetype');
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
         .t-12 {{ font-size: 1em; }}
        .t-14 {{ font-size: 1.1em; }}
        .t-16 {{ font-size: 1.5em; }}
        .t-18 {{ font-size: 1.7em; }}
        .t-20 {{ font-size: 1.9em; }}
        .t-22 {{ font-size: 2.1em; }}

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
            font-size: 1.4em;
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
    <div class='text-center t-16'><b>แบบสัญญา</b></div>
    <div class='text-center t-16'><b>สัญญาซื้อขาย</b></div>
    <div class='text-right t-14'>สัญญาเลขที่…{result.SPAContractNumber}.</div>
</br>
    <p class='tab2 t-14'>
        สัญญาฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.)
        ตำบล/แขวง ทุ่งสองห้อง อำเภอ/เขต หลักสี่ 
        จังหวัด กรุงเทพ เมื่อ {CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now)}
        ระหว่าง {result.Contract_Organization}
        โดย {result.SignatoryName}
        ซึ่งต่อไปในสัญญานี้เรียกว่า “ผู้ซื้อ” ฝ่ายหนึ่ง กับ {result.ContractorName}
    </p>
 
    {(result.ContractorType == "นิติบุคคล"
        ? $@"<p class='tab2 t-14'>
                ซึ่งจดทะเบียนเป็นนิติบุคคล ณ {result.ContractorName} มี
                สำนักงานใหญ่อยู่เลขที่ {result.ContractorAddressNo} ถนน {result.ContractorStreet} ตำบล/แขวง {result.ContractorSubDistrict}
                อำเภอ/เขต {result.ContractorDistrict} จังหวัด {result.ContractorProvince}
                โดย {result.ContractorSignatoryName}
                มีอำนาจลงนามผูกพันนิติบุคคลปรากฏตามหนังสือรับรองของสำนักงานทะเบียนหุ้นส่วนบริษัท
                ลงวันที่ {CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now)} (5)(และหนังสือมอบอำนาจลง {CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now)}) แนบท้ายสัญญานี้
            </p>"
        : $@"<p class='tab2 t-14'>
                (6)(ในกรณีที่ผู้รับจ้างเป็นบุคคลธรรมดาให้ใช้ข้อความว่า กับ {result.ContractorName}
                อยู่บ้านเลขที่ {result.ContractorAddressNo} ถนน {result.ContractorStreet} ตำบล/แขวง {result.ContractorSubDistrict}
                อำเภอ/เขต {result.ContractorDistrict} จังหวัด {result.ContractorProvince}
                ผู้ถือบัตรประจำตัวประชาชนเลขที่ {result.CitizenId} ดังปรากฏตามสำเนาบัตรประจำตัวประชาชนแนบท้ายสัญญานี้) ซึ่งต่อไปใน
                สัญญานี้เรียกว่า “ผู้ให้เช่า” อีกฝ่ายหนึ่ง
            </p>")}
</br>
    <p class='tab2 t-14'>คู่สัญญาได้ตกลงกันมีข้อความดังต่อไปนี้</p>

    <p class='tab2 t-14'><b>ข้อ ๑ ข้อตกลงซื้อขาย</b></p>
    <p class='tab3 t-14'>
        ผู้ซื้อตกลงซื้อและผู้ขายตกลงขาย {result.ProductDescription}
        จำนวน {result.Quantity} ({result.Unit}) เป็นราคาทั้งสิ้น {result.TotalAmount?.ToString("N2") ?? "0.00"} บาท ({CommonDAO.NumberToThaiText(result.TotalAmount ?? 0)})
        ซึ่งได้รวมภาษีมูลค่าเพิ่มจำนวน {result.VatAmount?.ToString("N2") ?? "0.00"} บาท ({CommonDAO.NumberToThaiText(result.VatAmount ?? 0)}) ตลอดจนภาษีอากรอื่นๆ และค่าใช้จ่ายทั้งปวงด้วยแล้ว
    </p>
    <p class='tab2 t-14'><b>ข้อ ๒ การรับรองคุณภาพ</b></p>
    <p class='tab3 t-14'>ผู้ขายรับรองว่าสิ่งของที่ขายให้ตามสัญญานี้เป็นของแท้ ของใหม่ ไม่เคยใช้งานมาก่อนไม่เป็น
ของเก่าเก็บ และมีคุณภาพและคุณสมบัติไม่ต่ำกว่าที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก ๑ ในกรณีที่
เป็นการซื้อสิ่งของซึ่งจะต้องมีการตรวจทดสอบ ผู้ขายรับรองว่า เมื่อตรวจทดสอบแล้วต้องมีคุณภาพและ
คุณสมบัติไม่ต่ำกว่าที่กำหนดไว้ตามสัญญานี้ด้วย
    </p>
    <p class='tab2 t-14'><b>ข้อ ๓ เอกสารอันเป็นส่วนหนึ่งของสัญญา</b></p>
    <p class='tab3 t-14'>เอกสารแนบท้ายสัญญาดังต่อไปนี้ให้ถือเป็นส่วนหนึ่งของสัญญานี้</p>
   {htmlDocAtt} 
    <p class='tab3 t-14'>ความใดในเอกสารแนบท้ายสัญญาที่ขัดหรือแย้งกับข้อความในสัญญานี้ ให้ใช้ข้อความใน
สัญญานี้บังคับ และในกรณีที่เอกสารแนบท้ายสัญญาขัดแย้งกันเอง ผู้ขายจะต้องปฏิบัติตามคำวินิจฉัยของ 
ผู้ซื้อ คำวินิจฉัยของผู้ซื้อให้ถือเป็นที่สุด และผู้ขายไม่มีสิทธิเรียกร้องราคา ค่าเสียหาย หรือค่าใช้จ่ายใดๆเพิ่มเติมจากผู้ซื้อทั้งสิ้น
    </p>
   <p class='tab2 t-14'><b>ข้อ ๔ การส่งมอบ</b></p>
    <p class='tab3 t-14'>
        ผู้ขายจะส่งมอบสิ่งของที่ซื้อขายตามสัญญาให้แก่ผู้ซื้อ ณ {result.DeliveryLocation} ภายใน {CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now)}
        ให้ถูกต้องและครบถ้วนตามที่กำหนดไว้ในข้อ ๑ แห่ง สัญญานี้ พร้อมทั้งหีบห่อหรือเครื่องรัดพันผูกโดยเรียบร้อย
    </p>
    <p class='tab3 t-14'>
        การส่งมอบสิ่งของตามสัญญานี้ ไม่ว่าจะเป็นการส่งมอบเพียงครั้งเดียว หรือส่งมอบหลายครั้ง
ผู้ขายจะต้องแจ้งกำหนดเวลาส่งมอบแต่ละครั้งโดยทำเป็นหนังสือนำไปยื่นต่อผู้ซื้อ ณ {result.DeliveryNotifyLocation}
        ในวันและเวลาทำการของผู้ซื้อ ก่อนวันส่งมอบไม่น้อยกว่า {result.DeliveryNotifyDays} วันทำการของผู้ซื้อ
    </p>
    <p class='tab2 t-14'><b>ข้อ ๕ การตรวจรับ</b></p>
    <p class='tab3 t-14'>เมื่อผู้ซื้อได้ตรวจรับสิ่งของที่ส่งมอบและเห็นว่าถูกต้องครบถ้วนตามสัญญาแล้ว ผู้ซื้อจะออก
หลักฐานการรับมอบเป็นหนังสือไว้ให้ เพื่อผู้ขายนำมาเป็นหลักฐานประกอบการขอรับเงิน ค่าสิ่งของนั้น</p>
    <p class='tab3 t-14'>ถ้าผลของการตรวจรับปรากฏว่า สิ่งของที่ผู้ขายส่งมอบไม่ตรงตามข้อ ๑ ผู้ซื้อทรงไว้ซึ่งสิทธิ
ที่จะไม่รับสิ่งของนั้นในกรณีเช่นว่านี้ ผู้ขายต้องรีบนำสิ่งของนั้นกลับคืนโดยเร็วที่สุดเท่าที่จะทำได้และนำ
สิ่งของมาส่งมอบให้ใหม่หรือต้องทำการแก้ไขให้ถูกต้องตามสัญญาด้วยค่าใช้จ่ายของผู้ขายเอง และระยะ
เวลาที่เสียไปเพราะเหตุดังกล่าวผู้ขายจะนำมาอ้างเป็นเหตุขอขยายเวลาส่งมอบตามสัญญาหรือ ของดหรือ
ลดค่าปรับไม่ได้ </p>
    <p class='tab3 t-14'>(๑๒) ในกรณีที่ผู้ขายส่งมอบสิ่งของถูกต้องแต่ไม่ครบจำนวน หรือส่งมอบครบจำนวน 
แต่ไม่ถูกต้องทั้งหมด ผู้ซื้อจะตรวจรับเฉพาะส่วนที่ถูกต้อง โดยออกหลักฐานการตรวจรับเฉพาะส่วนนั้น
ก็ได้ (ความในวรรคสามนี้จะไม่กำหนดไว้ในกรณีที่ผู้ซื้อต้องการสิ่งของทั้งหมดในคราวเดียวกัน 
หรือการซื้อสิ่งของที่ประกอบเป็นชุดหรือหน่วย ถ้าขาดส่วนประกอบอย่างหนึ่งอย่างใดไปแล้ว จะไม่สามารถ
ใช้งานได้ โดยสมบูรณ์)
    </p>
   <p class='tab2 t-14'><b>ข้อ ๖ การชำระเงิน</b></p>
<p class='tab3 t-14'>
    (๑๓ ก) ผู้ซื้อตกลงชำระเงินค่าสิ่งของตามข้อ ๑ ให้แก่ผู้ขาย เมื่อผู้ซื้อได้รับมอบสิ่งของตามข้อ ๕ ไว้โดยครบถ้วนแล้ว
</p>
<p class='tab3 t-14'>
    (๑๓ ข) ผู้ซื้อตกลงชำระเงินค่าสิ่งของตามข้อ ๑ ให้แก่ผู้ขาย ดังนี้
</p>
<p class='tab3 t-14'>
    ๖.๑ เงินล่วงหน้า จำนวน {result.AdvancePayment} บาท ({CommonDAO.NumberToThaiText(result.AdvancePayment ?? 0)}) จะจ่ายให้ภายใน {result.PaymentDueDays} 
วัน นับถัดจากวันลงนาม
ในสัญญา ทั้งนี้ โดยผู้ขายจะต้องนำหลักประกันเงินล่วงหน้าเป็น {result.PaymentGuaranteeType} (หนังสือค้ำประกันหรือหนังสือค้ำประกัน
อิเล็กทรอนิกส์ของธนาคารภายในประเทศหรือพันธบัตรรัฐบาลไทย){result.PaymentGuaranteeType}เต็มตามจำนวนเงินล่วงหน้าที่จะได้รับ มามอบให้แก่ผู้ซื้อเป็นหลักประกันการชำระคืนเงินล่วงหน้าก่อนการรับชำระเงินล่วงหน้านั้น และผู้ซื้อจะคืนหลักประกันเงินล่วงหน้าให้แก่ผู้ขายเมื่อผู้ซื้อจ่ายเงินที่เหลือตามข้อ ๖.๒
</p>
<p class='tab3 t-14'>
    ๖.๒ เงินที่เหลือ จำนวน {result.AdvancePayment} บาท ({CommonDAO.NumberToThaiText(result.AdvancePayment ?? 0)}) 
จะจ่ายให้เมื่อผู้ซื้อได้รับมอบสิ่งของ ตามข้อ ๕ ไว้โดยถูกต้องครบถ้วนแล้ว
</p>
<p class='tab3 t-14'>
    (๑๔) การจ่ายเงินตามเงื่อนไขแห่งสัญญานี้ ผู้ซื้อจะโอนเงินเข้าบัญชีเงินฝากธนาคารของผู้ขาย ชื่อธนาคาร {result.SaleBankName} สาขา {result.SaleBankBranch} ชื่อบัญชี {result.SaleBankAccountName} 
เลขที่บัญชี {result.SaleBankAccountNumber} ทั้งนี้ ผู้ขายตกลงเป็นผู้รับภาระเงินค่าธรรมเนียม หรือค่าบริการอื่นใดเกี่ยวกับการ
โอน รวมทั้งค่าใช้จ่ายใดๆ (ถ้ามี) ที่ธนาคารเรียกเก็บ และยินยอมให้มีการหักเงินดังกล่าวจากจำนวนเงินโอนใน
งวดนั้นๆ
</p>

<p class='tab2 t-14'><b>ข้อ ๗ การรับประกันความชำรุดบกพร่อง</b></p>
<p class='tab3 t-14'>
    ผู้ขายตกลงรับประกันความชำรุดบกพร่องหรือขัดข้องของสิ่งของตามสัญญานี้ เป็นเวลา{result.WarrantyPeriodYears} ปี {result.WarrantyPeriodMonths} เดือน
นับถัดจากวันที่ผู้ซื้อได้รับมอบสิ่งของทั้งหมดไว้โดยถูกต้องครบถ้วนตามสัญญา โดยภายในกำหนด
เวลาดังกล่าว หากสิ่งของ ตามสัญญานี้เกิดชำรุดบกพร่องหรือขัดข้องอันเนื่องมาจากการใช้งานตามปกติผู้ขาย
จะต้องจัดการซ่อมแซมหรือแก้ไขให้อยู่ในสภาพที่ใช้การได้ดีดังเดิม ภายใน {result.DaysToRepairAfterNoti} วัน นับถัดจากวันที่ได้รับแจ้ง
จากผู้ซื้อ โดยไม่คิดค่าใช้จ่ายใดๆ ทั้งสิ้น หากผู้ขายไม่จัดการซ่อมแซมหรือแก้ไขภายในกำหนดเวลาดังกล่าว
ผู้ซื้อมีสิทธิที่จะทำการนั้นเองหรือจ้างผู้อื่นให้ทำการนั้นแทนผู้ขาย โดยผู้ขายต้องเป็นผู้ออกค่าใช้จ่ายเองทั้งสิ้น
</p>
<p class='tab3 t-14'>
    ในกรณีเร่งด่วนจำเป็นต้องรีบแก้ไขเหตุชำรุดบกพร่องหรือขัดข้องโดยเร็ว และไม่อาจรอคอย
ให้ผู้ขายแก้ไขในระยะเวลาที่กำหนดไว้ตามวรรคหนึ่งได้ ผู้ซื้อมีสิทธิเข้าจัดการแก้ไขเหตุชำรุดบกพร่องหรือ
ขัดข้องนั้นเอง หรือให้ผู้อื่นแก้ไขความชำรุดบกพร่องหรือขัดข้อง โดยผู้ขายต้องรับผิดชอบชำระค่าใช้จ่ายทั้งหมด
</p>
<p class='tab3 t-14'>
    การที่ผู้ซื้อทำการนั้นเอง หรือให้ผู้อื่นทำการนั้นแทนผู้ขาย ไม่ทำให้ผู้ขายหลุดพ้นจากความ
รับผิดตามสัญญา หากผู้ขายไม่ชดใช้ค่าใช้จ่ายหรือค่าเสียหายตามที่ผู้ซื้อเรียกร้องผู้ซื้อมีสิทธิบังคับจากหลักประกันการปฏิบัติตามสัญญาได้
</p>

<p class='tab2 t-14'><b>ข้อ ๘ หลักประกันการปฏิบัติตามสัญญา</b></p>
<p class='tab3 t-14'>
    ในขณะทำสัญญานี้ผู้ขายได้นำหลักประกันเป็น {result.GuaranteeType} เป็นจำนวนเงิน {result.GuaranteeAmount} บาท 
({CommonDAO.NumberToThaiText(result.GuaranteeAmount ?? 0)}) ซึ่งเท่ากับร้อยละ {result.GuaranteePercent} ของราคาทั้งหมดตามสัญญามามอบให้แก่ผู้ซื้อเพื่อเป็นหลัก
ประกันการปฏิบัติตามสัญญานี้ (๑๘) กรณีผู้ขายใช้หนังสือค้ำประกันมาเป็นหลักประกันการปฏิบัติตามสัญญา 
หนังสือค้ำประกันดังกล่าวจะต้องออกโดยธนาคารที่ประกอบกิจการในประเทศไทย หรือโดยบริษัทเงินทุนหรือ
บริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกัน
ตามประกาศของธนาคารแห่งประเทศไทยตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบ
ตามแบบที่คณะกรรมการนโยบายการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐกำหนด หรืออาจเป็นหนังสือค้ำ
ประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้ และจะต้องมีอายุการค้ำประกันตลอดไปจนกว่า
ผู้ขายพ้นข้อผูกพันตามสัญญานี้หลักประกันที่ผู้ขายนำมามอบให้ตามวรรคหนึ่ง จะต้องมีอายุครอบคลุมความ
รับผิดทั้งปวงของผู้ขายตลอดอายุสัญญานี้ ถ้าหลักประกันที่ผู้ขายนำมามอบให้ดังกล่าวลดลงหรือเสื่อมค่าลง 
หรือมีอายุไม่ครอบคลุมถึงความรับผิดของผู้ขายตลอดอายุสัญญา ไม่ว่าด้วยเหตุใดๆ ก็ตาม รวมถึงกรณีผู้ขาย
ส่งมอบสิ่งของล่าช้าเป็นเหตุให้ระยะเวลาส่งมอบหรือวันครบกำหนดความรับผิดในความชำรุดบกพร่องตาม
สัญญาเปลี่ยนแปลงไปไม่ว่าจะเกิดขึ้นคราวใดผู้ขายต้องหาหลักประกันใหม่หรือหลักประกันเพิ่มเติมให้มี
จำนวนครบถ้วนตามวรรคหนึ่งมามอบให้แก่ผู้ซื้อภายใน {result.NewGuaranteeDays} วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือ
จากผู้ซื้อหลักประกันที่ผู้ขายนำมามอบไว้ตามข้อนี้ ผู้ซื้อจะคืนให้แก่ผู้ขายโดยไม่มีดอกเบี้ยเมื่อผู้ขายพ้นจากข้อ
ผูกพันและความรับผิดทั้งปวงตามสัญญานี้แล้ว

</p>

<p class='tab2 t-14'><b>ข้อ ๙ การบอกเลิกสัญญา</b></p>
<p class='tab3 t-14'>ถ้าผู้ขายไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่ง หรือเมื่อครบกำหนดส่งมอบสิ่งของตามสัญญานี้
แล้ว หากผู้ขายไม่ส่งมอบสิ่งของที่ตกลงขายให้แก่ผู้ซื้อหรือส่งมอบไม่ถูกต้อง หรือไม่ครบจำนวน ผู้ซื้อมีสิทธิ
บอกเลิกสัญญาทั้งหมดหรือแต่บางส่วนได้ การใช้สิทธิบอกเลิกสัญญานั้นไม่กระทบสิทธิของผู้ซื้อที่จะเรียก
ร้องค่าเสียหายจากผู้ขายในกรณีที่ผู้ซื้อใช้สิทธิบอกเลิกสัญญา ผู้ซื้อมีสิทธิริบหรือบังคับจากหลักประกันตาม
(๑๙)(ข้อ ๖ และ) ข้อ ๘ เป็นจำนวนเงินทั้งหมดหรือแต่บางส่วนก็ได้ แล้วแต่ผู้ซื้อจะเห็นสมควรและถ้าผู้ซื้อ
จัดซื้อสิ่งของจากบุคคลอื่นเต็มจำนวนหรือเฉพาะจำนวนที่ขาดส่ง แล้วแต่กรณี ภายในกำหนด {result.TerminationNewMonths} เดือน
นับถัดจากวันบอกเลิกสัญญา ผู้ขายจะต้องชดใช้ราคาที่เพิ่มขึ้นจากราคาที่กำหนดไว้ในสัญญานี้ด้วย
</p>

<p class='tab2 t-14'><b>ข้อ ๑๐ ค่าปรับ</b></p>
<p class='tab3 t-14'>
    ในกรณีที่ผู้ซื้อมิได้ใช้สิทธิบอกเลิกสัญญาตามข้อ ๙ ผู้ขายจะต้องชำระค่าปรับให้ผู้ซื้อเป็น
รายวันในอัตราร้อยละ {result.FineRatePerDay}(%) ของราคาสิ่งของที่ยังไม่ได้รับมอบ นับถัดจากวันครบกำหนดตามสัญญา
จนถึงวันที่ผู้ขายได้นำสิ่งของมาส่งมอบให้แก่ผู้ซื้อจนถูกต้องครบถ้วนตามสัญญาการคิดค่าปรับในกรณีสิ่ง
ของที่ตกลงซื้อขายประกอบกันเป็นชุด แต่ผู้ขายส่งมอบเพียงบางส่วน หรือขาดส่วนประกอบส่วนหนึ่งส่วนใด
ไปทำให้ไม่สามารถใช้การได้โดยสมบูรณ์ ให้ถือว่า ยังไม่ได้ส่งมอบสิ่งของนั้นเลยและให้คิดค่าปรับจากราคา
สิ่งของเต็มทั้งชุดในระหว่างที่ผู้ซื้อยังมิได้ใช้สิทธิบอกเลิกสัญญานั้น หากผู้ซื้อเห็นว่าผู้ขายไม่อาจปฏิบัติตาม
สัญญาต่อไปได้ ผู้ซื้อจะใช้สิทธิบอกเลิกสัญญาและริบหรือบังคับจากหลักประกันตาม (๒๒)(ข้อ ๖ และ) ข้อ ๘ 
กับเรียกร้องให้ชดใช้ราคาที่เพิ่มขึ้นตามที่กำหนดไว้ในข้อ ๙ วรรคสองก็ได้และถ้าผู้ซื้อได้แจ้งข้อเรียกร้อง
ให้ชำระค่าปรับไปยังผู้ขายเมื่อครบกำหนดส่งมอบแล้ว ผู้ซื้อมีสิทธิที่จะปรับผู้ขายจนถึงวันบอกเลิกสัญญาได้อีกด้วย
</p>

<p class='tab2 t-14'><b>ข้อ ๑๑ การบังคับค่าปรับ ค่าเสียหาย และค่าใช้จ่าย</b></p>
<p class='tab3 t-14'>ในกรณีที่ผู้ขายไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่งด้วยเหตุใดๆ ก็ตาม จนเป็นเหตุให้เกิดค่าปรับ
ค่าเสียหาย หรือค่าใช้จ่ายแก่ผู้ซื้อ ผู้ขายต้องชดใช้ค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายดังกล่าวให้แก่ผู้ซื้อโดย
สิ้นเชิงภายในกำหนด {result.FinePeriodAfterNotiDays} วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ซื้อ หากผู้ขายไม่ชดใช้ให้ถูกต้อง
ครบถ้วนภายในระยะเวลาดังกล่าวให้ผู้ซื้อมีสิทธิที่จะหักเอาจากจำนวนเงินค่าสิ่งของที่ซื้อขายที่ต้องชำระหรือ
บังคับจากหลักประกันการปฏิบัติตามสัญญาได้ทันทีหากค่าปรับ ค่าเสียหายหรือค่าใช้จ่ายที่บังคับจากเงินค่าสิ่ง
ของที่ซื้อขายที่ต้องชำระ หรือหลักประกันการปฏิบัติตามสัญญาแล้วยังไม่เพียงพอ ผู้ขายยินยอมชำระส่วนที่
เหลือที่ยังขาดอยู่จนครบถ้วนตามจำนวนค่าปรับค่าเสียหายหรือค่าใช้จ่ายนั้น ภายในกำหนด {result.DefectFinePeriodDays} วัน 
นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ซื้อ หากมีเงินค่าสิ่งของที่ซื้อขายตามสัญญาที่หักไว้จ่ายเป็นค่าปรับ
ค่าเสียหาย หรือค่าใช้จ่ายแล้วยังเหลืออยู่อีกเท่าใด ผู้ซื้อจะคืนให้แก่ผู้ขายทั้งหมด
</p>

<p class='tab2 t-14'><b>ข้อ ๑๒ การงดหรือลดค่าปรับ หรือขยายเวลาส่งมอบ</b></p>
<p class='tab3 t-14'>
    ในกรณีที่มีเหตุเกิดจากความผิดหรือความบกพร่องของฝ่ายผู้ซื้อ หรือเหตุสุดวิสัย หรือเกิดจาก
พฤติการณ์อันหนึ่งอันใดที่ผู้ขายไม่ต้องรับผิดตามกฎหมาย หรือเหตุอื่นตามที่กำหนดในกฎกระทรวง ซึ่งออก
ตามความในกฎหมายว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ ทำให้ผู้ขายไม่สามารถส่งมอบสิ่งของ
ตามเงื่อนไขและกำหนดเวลาแห่งสัญญานี้ได้ ผู้ขายมีสิทธิของดหรือลดค่าปรับหรือขยายเวลาส่งมอบตามสัญญา
ได้ โดยจะต้องแจ้งเหตุหรือพฤติการณ์ดังกล่าวพร้อมหลักฐานเป็นหนังสือให้ผู้ซื้อทราบภายใน ๑๕ (สิบห้า) วัน
 นับถัดจากวันที่เหตุนั้นสิ้นสุดลง หรือตามที่กำหนดในกฎกระทรวงดังกล่าวถ้าผู้ขายไม่ปฏิบัติให้เป็นไปตาม
ความในวรรคหนึ่ง ให้ถือว่าผู้ขายได้สละสิทธิเรียกร้องในการที่จะของดหรือลดค่าปรับหรือขยายเวลาส่งมอบ
ตามสัญญา โดยไม่มีเงื่อนไขใดๆ ทั้งสิ้น เว้นแต่กรณีเหตุเกิดจากความผิดหรือความบกพร่องของฝ่ายผู้ซื้อ
ซึ่งมีหลักฐานชัดแจ้งหรือผู้ซื้อทราบดีอยู่แล้วตั้งแต่ต้นการงดหรือลดค่าปรับหรือขยายเวลาส่งมอบตามสัญญา
ตามวรรคหนึ่ง อยู่ในดุลพินิจของผู้ซื้อที่จะพิจารณาตามที่เห็นสมควร
</p>

<p class='tab2 t-14'><b>ข้อ ๑๓ การใช้เรือไทย</b></p>
<p class='tab3 t-14'>ถ้าสิ่งของที่จะต้องส่งมอบให้แก่ผู้ซื้อตามสัญญานี้ เป็นสิ่งของที่ผู้ขายจะต้องสั่งหรือนำเข้า
มาจากต่างประเทศ และสิ่งของนั้นต้องนำเข้ามาโดยทางเรือในเส้นทางเดินเรือที่มีเรือไทยเดินอยู่ และสามารถ
ให้บริการรับขนได้ตามที่รัฐมนตรีว่าการกระทรวงคมนาคมประกาศกำหนด ผู้ขายต้องจัดการให้สิ่งของดังกล่าว
บรรทุกโดยเรือไทยหรือเรือที่มีสิทธิเช่นเดียวกับเรือไทยจากต่างประเทศมายังประเทศไทย เว้นแต่จะได้รับ
อนุญาตจากกรมเจ้าท่าก่อนบรรทุกของนั้นลงเรืออื่นที่มิใช่เรือไทยหรือเป็นของที่รัฐมนตรี ว่าการกระทรวง
คมนาคมประกาศยกเว้นให้บรรทุกโดยเรืออื่นได้ ทั้งนี้ ไม่ว่าการสั่งหรือนำเข้าสิ่งของดังกล่าวจากต่างประเทศ
จะเป็นแบบใด
</p>
<p class='tab3 t-14'>
    ในการส่งมอบสิ่งของตามสัญญาให้แก่ผู้ซื้อ ถ้าสิ่งของนั้นเป็นสิ่งของตามวรรคหนึ่ง ผู้ขายจะ
ต้องส่งมอบใบตราส่ง(Bill of Lading) หรือสำเนาใบตราส่งสำหรับของนั้น ซึ่งแสดงว่าได้บรรทุกมาโดย
เรือไทยหรือเรือที่มีสิทธิเช่นเดียวกับเรือไทยให้แก่ผู้ซื้อพร้อมกับการส่งมอบสิ่งของด้วย
</p>
<p class='tab3 t-14'>
    ในกรณีที่สิ่งของดังกล่าวไม่ได้บรรทุกจากต่างประเทศมายังประเทศไทย โดยเรือไทยหรือเรือที่
มีสิทธิเช่นเดียวกับเรือไทยผู้ขายต้องส่งมอบหลักฐานซึ่งแสดงว่าได้รับอนุญาตจากกรมเจ้าท่าให้บรรทุกของโดย
เรืออื่นได้หรือหลักฐานซึ่งแสดงว่าได้ชำระค่าธรรมเนียมพิเศษเนื่องจากการไม่บรรทุกของโดยเรือไทยตามกฎหมายว่าด้วยการส่งเสริมการพาณิชยนาวีแล้วอย่างใดอย่างหนึ่งแก่ผู้ซื้อด้วย
</p>
<p class='tab3 t-14'>
    ในกรณีที่ผู้ขายไม่ส่งมอบหลักฐานอย่างใดอย่างหนึ่งดังกล่าวในวรรคสองและวรรคสามให้แก่ผู้ซื้อ แต่จะขอส่งมอบสิ่งของดังกล่าวให้ผู้ซื้อก่อนโดยยังไม่รับชำระเงินค่าสิ่งของ ผู้ซื้อมีสิทธิรับสิ่งของดังกล่าวไว้ก่อนและชำระเงินค่าสิ่งของเมื่อผู้ขายได้ปฏิบัติถูกต้องครบถ้วนดังกล่าวแล้วได้
</p>
<p class='tab3 t-14'>
    สัญญานี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน คู่สัญญาได้อ่านและเข้าใจข้อความ โดยละเอียดตลอดแล้ว จึงได้ลงลายมือชื่อพร้อมทั้งประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยาน และคู่สัญญาต่างยึดถือไว้ฝ่ายละหนึ่งฉบับ
</p>

<!-- Signature Block -->


</br>
</br>
{signatoryTableHtml}
    <!-- ... -->
    <!-- Add all other clauses and signature blocks as in your Word logic -->
<div style='page-break-before: always;'></div>
<p class='text-center t-16'><b>วิธีปฏิบัติเกี่ยวกับสัญญาซื้อขายคอมพิวเตอร์</b></p>
<p class='tab2 t-14'>(๑) ให้ระบุเลขที่สัญญาในปีงบประมาณหนึ่งๆ ตามลำดับ</p>
<p class='tab2 t-14'>(๒) ให้ระบุชื่อของหน่วยงานของรัฐที่เป็นนิติบุคคล เช่น กรม ก. หรือรัฐวิสาหกิจ ข. เป็นต้น</p>
<p class='tab2 t-14'>(๓) ให้ระบุชื่อและตำแหน่งของหัวหน้าหน่วยงานของรัฐที่เป็นนิติบุคคลนั้น หรือผู้ที่ได้รับมอบอำนาจ เช่น นาย ก. อธิบดีกรม……………… หรือ นาย ข. ผู้ได้รับมอบอำนาจจากอธิบดีกรม………………..</p>
<p class='tab2 t-14'>(๔) ให้ระบุชื่อผู้ขาย</p>

<p class='tab3 t-14'>ก. กรณีนิติบุคคล เช่น ห้างหุ้นส่วนสามัญจดทะเบียน ห้างหุ้นส่วนจำกัด บริษัทจำกัด</p>
<p class='tab3 t-14'>ข. กรณีบุคคลธรรมดา ให้ระบุชื่อและที่อยู่</p>
<p class='tab2 t-14'>(๕) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
<p class='tab2 t-14'>(๖) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
<p class='tab2 t-14'>(๗) ให้ระบุว่าเป็นการซื้อสิ่งของตามตัวอย่าง หรือรายการละเอียด หรือแค็ตตาล็อก หรือแบบรูปรายการ หรืออื่นๆ (ให้ระบุ) และปกติจะต้องกำหนดไว้ด้วยว่าสิ่งของที่จะซื้อนั้น เป็นของแท้ เป็นของใหม่ ไม่เคยใช้งานมาก่อน </p>
<p class='tab2 t-14'>(๘) ให้ระบุหน่วยที่ใช้ เช่น กิโลกรัม ชิ้น เมตร เป็นต้น</p>
<p class='tab2 t-14'>(๙) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
<p class='tab2 t-14'>(๑๐) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
<p class='tab2 t-14'>(๑๑) กำหนดเวลาส่งมอบจะต้องแจ้งล่วงหน้าไม่น้อยกว่ากี่วัน ให้อยู่ในดุลพินิจของผู้ซื้อโดยตกลงกับผู้ขาย โดยปกติควรจะกำหนดไว้ประมาณ ๓ วันทำการ เพื่อที่ผู้ซื้อจะได้จัดเตรียมเจ้าหน้าที่ไว้ตรวจรับของนั้น</p>
<p class='tab2 t-14'>ในกรณีที่มีการส่งมอบสิ่งของหลายครั้ง ให้ระบุวันเวลาที่ส่งมอบแต่ละครั้งไว้ด้วย และในกรณีที่มีการติดตั้งด้วย ให้แยกกำหนดเวลาส่งมอบ และกำหนดเวลาการติดตั้งออกจากกัน</p>
<p class='tab2 t-14'>(๑๒) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
<p class='tab2 t-14'>(๑๓) ให้หน่วยงานของรัฐเลือกใช้ตามความเหมาะสม</p>
<p class='tab3 t-14'>ข้อความในข้อ ๖ กรณีไม่มีการจ่ายเงินล่วงหน้าให้ผู้ขาย ให้เลือกใช้ข้อความในข้อ (๑๓ ก)</p>
<p class='tab3 t-14'>ข้อความในข้อ ๖ กรณีมีการจ่ายเงินล่วงหน้าให้ผู้ขาย ให้เลือกใช้ข้อความในข้อ (๑๓ ข)</p>
<p class='tab2 t-14'>(๑๔) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
<p class='tab2 t-14'>(๑๕) ระยะเวลารับประกันและระยะเวลาแก้ไขซ่อมแซมจะกำหนดเท่าใด แล้วแต่ลักษณะของสิ่งของที่ซื้อขายกัน โดยให้อยู่ในดุลพินิจของผู้ซื้อ เช่น เครื่องคำนวณไฟฟ้า กำหนดเวลารับประกัน ๑ ปี กำหนดเวลาแก้ไขภายใน ๗ วัน เป็นต้น ทั้งนี้ จะต้องประกาศให้ทราบในเอกสารเชิญชวนด้วย</p>
<p class='tab2 t-14'>(๑๖) “หลักประกัน” หมายถึง หลักประกันที่ผู้ขายนำมามอบไว้แก่หน่วยงานของรัฐ เมื่อลงนามในสัญญา เพื่อเป็นการประกันความเสียหายที่อาจจะเกิดขึ้นจากการปฏิบัติตามสัญญา ดังนี้</p>
<p class='tab3 t-14'>(๑)เงินสด </p>
<p class='tab3 t-14'>(๒)เช็คหรือดราฟท์ที่ธนาคารเซ็นสั่งจ่าย ซึ่งเป็นเช็คหรือดราฟท์ลงวันที่ที่ใช้เช็คหรือดราฟท์นั้นชำระต่อเจ้าหน้าที่ หรือก่อนวันนั้นไม่เกิน ๓ วันทำการ </p>
<p class='tab3 t-14'>(๓)หนังสือค้ำประกันของธนาคารภายในประเทศตามตัวอย่างที่คณะกรรมการนโยบายกำหนด โดยอาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้</p>
<p class='tab3 t-14'>(๔)หนังสือค้ำประกันของบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบ โดยอนุโลมให้ใช้ตามตัวอย่างหนังสือ ค้ำประกันของธนาคารที่คณะกรรมการนโยบายกำหนด</p>
<p class='tab3 t-14'>(๕)พันธบัตรรัฐบาลไทย</p>
<p class='tab2 t-14'>(๑๗) ให้กำหนดจำนวนเงินหลักประกันการปฏิบัติตามสัญญาตามระเบียบกระทรวงการคลังว่าด้วยหลักเกณฑ์การจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ ข้อ ๑๖๘</p>
<p class='tab2 t-14'>(๑๘) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
<p class='tab2 t-14'>(๑๙) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
<p class='tab2 t-14'>(๒๐) กำหนดเวลาที่ผู้ซื้อจะซื้อสิ่งของจากแหล่งอื่นเมื่อบอกเลิกสัญญาและมีสิทธิเรียกเงินในส่วนที่เพิ่มขึ้นจากราคาที่กำหนดไว้ในสัญญานั้น ให้อยู่ในดุลพินิจของผู้ซื้อโดยตกลงกับผู้ขาย และโดยปกติแล้วไม่ควรเกิน ๓ เดือน</p>
<p class='tab2 t-14'>(๒๑) อัตราค่าปรับตามสัญญาข้อ 10 ให้กำหนดเป็นรายวันในอัตราระหว่างร้อยละ ๐.๑๐-๐.๒๐ ตามระเบียบกระทรวงการคลังว่าด้วยหลักเกณฑ์การจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ข้อ ๑๖๒ ส่วนกรณีจะปรับร้อยละเท่าใด ให้อยู่ในดุลพินิจของหน่วยงานของรัฐผู้ซื้อที่จะพิจารณา โดยคำนึงถึงราคาและลักษณะของพัสดุที่ซื้อ ซึ่งอาจมีผลกระทบต่อการที่ผู้ขายจะหลีกเลี่ยงไม่ปฏิบัติตามสัญญา แต่ทั้งนี้การที่จะกำหนดค่าปรับเป็นร้อยละเท่าใด จะต้องกำหนดไว้ในเอกสารเชิญชวนด้วย</p>
<p class='tab2 t-14'>(๒๒) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p> 




</body>
</html>
";

  
       
        return html;


    }
    #endregion 4.1.1.2.11.สัญญาเช่าคอมพิวเตอร์ ร.309-60

}
