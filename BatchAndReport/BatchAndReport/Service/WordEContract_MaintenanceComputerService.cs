using BatchAndReport.DAO;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Threading.Tasks;


public class WordEContract_MaintenanceComputerService
{
    private readonly WordServiceSetting _w;
    private readonly  Econtract_Report_SMCDAO _econtractReportSMCDAO;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
    private readonly EContractDAO _eContractDAO;
    private readonly E_ContractReportDAO _eContractReportDAO;
    public WordEContract_MaintenanceComputerService(WordServiceSetting ws
        , Econtract_Report_SMCDAO econtractReportSMCDAO
             , IConverter pdfConverter
        ,
EContractDAO eContractDAO
        ,
E_ContractReportDAO eContractReportDAO

        )
    {
        _w = ws;
        _econtractReportSMCDAO = econtractReportSMCDAO;
        _pdfConverter = pdfConverter; // กำหนดค่า PDF Converter
        _eContractDAO = eContractDAO;
        _eContractReportDAO = eContractReportDAO;
    }
    #region 4.1.1.2.12.สัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ร.310-60
    public async Task<byte[]> OnGetWordContact_MaintenanceComputer(string id)
    {
        var result =await _econtractReportSMCDAO.GetSMCAsync(id);
        if (result == null)
        {
            throw new Exception("SMC data not found.");
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
                body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("สัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์", "000000", "36"));
                // 2. Document title and subtitle
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.RightParagraph("สัญญาเลขที่………….…… (1)...........……..……..."));



                
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
                    "สัญญานี้เรียกว่า “ผู้ซื้อ” อีกฝ่ายหนึ่ง กับ…" + result.ContractorName + "", null, "32"));

                }

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("คู่สัญญาได้ตกลงกันมีข้อความดังต่อไปนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ1 ขอบเขตของงาน", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ว่าจ้างตกลงจ้างและผู้รับจ้างตกลงรับจ้างให้บริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ อุปกรณ์การประมวลผล และระบบคอมพิวเตอร์ ตามเอกสารแนบท้ายสัญญาผนวก ๑ ซึ่งต่อไป   ในสัญญานี้เรียกว่า “คอมพิวเตอร์” ซึ่งติดตั้งอยู่ ณ "+result.Contract_Sign_Address+" ", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้รับจ้างจะจัดหาวัสดุสิ่งของชนิดที่ดีได้มาตรฐาน ใช้เครื่องมือดี และช่างผู้ชำนาญและฝีมือดีเพื่อใช้ในงานจ้างที่จำเป็นสำหรับการปฏิบัติงานตามสัญญา", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ๒ เอกสารอันเป็นส่วนหนึ่งของสัญญา", null, "32", true));

                // ไม่เจอ
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เอกสารแนบท้ายสัญญาดังต่อไปนี้ให้ถือว่าเป็นส่วนหนึ่งของสัญญานี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๒.๑ผนวก ๑ รายการคอมพิวเตอร์ที่บำรุงรักษาตามสัญญาจำนวน………(………) หน้า", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๒.๒ผนวก ๒ การกำหนดตัวถ่วงของคอมพิวเตอร์จำนวน………(………) หน้า", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("๒.๓ผนวก ๓ คาบเวลาที่ต้องการบำรุงรักษาจำนวน………(………) หน้า", null, "32"));
                // ไม่เจอ

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs(" และอัตราค่าบำรุงรักษา", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ความใดในเอกสารแนบท้ายสัญญาที่ขัดหรือแย้งกับข้อความในสัญญานี้ ให้ใช้ข้อความในสัญญานี้บังคับ และในกรณีที่เอกสารแนบท้ายสัญญาขัดแย้งกันเอง ผู้รับจ้างจะต้องปฏิบัติตามคำวินิจฉัยของ ผู้ว่าจ้าง คำวินิจฉัยของผู้ว่าจ้างให้ถือเป็นที่สุด และผู้รับจ้างไม่มีสิทธิเรียกร้องค่าจ้าง ค่าเสียหาย หรือค่าใช้จ่ายใดๆ เพิ่มเติมจากผู้ว่าจ้างทั้งสิ้น", null, "32"));

                string strServiceStartDate = CommonDAO.ToThaiDateStringCovert(result.ServiceStartDate ?? DateTime.Now);
                string strServiceEndDate = CommonDAO.ToThaiDateStringCovert(result.ServiceStartDate ?? DateTime.Now);

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ๓ ระยะเวลาให้บริการ", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้รับจ้างตกลงให้บริการตามสัญญานี้ ตั้งแต่วันที่ "+ strServiceStartDate + " ถึง "+ strServiceEndDate + " รวมเป็นเวลาทั้งสิ้น "+result.ServiceTotalYears+" ปี "+result.ServiceTotalMonths+" เดือน "+result.ServiceTotalDays+" วัน", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ๔ tค่าจ้างและการจ่ายเงิน", null, "32", true));

                string strServiceFee = CommonDAO.NumberToThaiText(result.ServiceFee ?? 0);
                string strVatAmount = CommonDAO.NumberToThaiText(result.VatAmount ?? 0);
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ว่าจ้างตกลงชำระค่าจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์เป็นเงินทั้งสิ้น "+result.ServiceFee+" บาท ("+ strServiceFee + ") ซึ่งได้รวมภาษีมูลค่าเพิ่มจำนวน "+result.VatAmount+" บาท ("+ strVatAmount + ") ตลอดจนค่าแรงงานค่าสิ่งของตลอดอายุสัญญา ภาษีอากรอื่น และค่าใช้จ่ายทั้งปวงไว้ด้วยแล้ว โดยผู้ว่าจ้างจะ  แบ่งจ่ายให้แก่ผู้รับจ้างเป็นงวดๆ รวม "+result.PaymentInstallment+" งวด ดังนี้", null, "32"));

                var installmentList =await _econtractReportSMCDAO.GetSMCInstallmentAsync(id);
                if (installmentList !=null && installmentList.ToList().Count>0) 
                {
                    foreach (var item in installmentList) 
                    {
                        body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("งวดที่ "+item.PayRound+" เป็นเงิน บาท "+item.TotalAmount+" จะจ่ายเมื่อผู้รับจ้างได้ดำเนินการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์เป็นเวลา "+item.RepairMonth+" เดือน และผู้ว่าจ้างได้ตรวจรับมอบงานตามสัญญาแล้ว", null, "32"));

                    }

                }            
                
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ค่าจ้างตามสัญญานี้ เป็นอัตราที่กำหนดไว้สำหรับการให้บริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ในเวลาตามรายละเอียดที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก ๓", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่มีการเปลี่ยนแปลงรายการตามเอกสารแนบท้ายสัญญาผนวก ๓ หรือมีการเปลี่ยนแปลงลักษณะเฉพาะของคอมพิวเตอร์ส่วนใดส่วนหนึ่งอันเป็นผลให้ต้องมีการเปลี่ยนแปลงแก้ไขอัตราค่าจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขตามที่ระบุไว้ในเอกสารแนบท้ายสัญญาผนวก ๓ ผู้ว่าจ้างหรือผู้รับจ้างมีสิทธิขอเปลี่ยนแปลงแก้ไขอัตราค่าจ้างบริการดังกล่าวได้ การเปลี่ยนแปลงแก้ไขอัตราค่าจ้างบริการดังกล่าวจะมีผลบังคับต่อเมื่อได้ระบุไว้ในผนวกเพิ่มเติม ซึ่งจะถือว่าเป็นส่วนหนึ่งแห่งสัญญานี้", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ๕ การรับประกันผลงาน", null, "32", true));

                string strPenaltyPerHours = CommonDAO.NumberToThaiText(result.PenaltyPerHours ?? 0);
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้รับจ้างตกลงบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ตามสัญญานี้ให้อยู่ในสภาพใช้งานได้ดีอยู่เสมอ โดยให้มีเวลาคอมพิวเตอร์ขัดข้องรวมตามเกณฑ์การคำนวณเวลาขัดข้อง ไม่เกินเดือนละ "+result.MaximumDownTimeHours+" ชั่วโมง หรือร้อยละ "+ result.MaximumDownPercents+ " ของเวลาใช้งานทั้งหมดของคอมพิวเตอร์ของเดือนนั้น แล้วแต่ตัวเลขใดจะมากกว่ากัน" +
                    "มิฉะนั้นผู้รับจ้างต้องยอมให้ผู้ว่าจ้างคิดค่าปรับเป็นรายชั่วโมง ในอัตราชั่วโมงละ "+result.PenaltyPerHours+" บาท ("+ strPenaltyPerHours + ") ในช่วงเวลาที่ไม่สามารถใช้คอมพิวเตอร์ได้ในส่วนที่เกินกว่ากำหนดเวลาขัดข้องข้างต้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เกณฑ์การคำนวณเวลาขัดข้องของคอมพิวเตอร์ตามวรรคหนึ่งให้เป็นไปดังนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("- กรณีที่คอมพิวเตอร์เกิดขัดข้องพร้อมกันหลายหน่วย ให้นับเวลาขัดข้องของหน่วยที่มี ตัวถ่วงมากที่สุดเพียงหน่วยเดียว", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("- กรณีความเสียหายอันสืบเนื่องมาจากความขัดข้องของคอมพิวเตอร์แตกต่างกัน เวลาที่ใช้ในการคำนวณค่าปรับจะเท่ากับเวลาขัดข้องของคอมพิวเตอร์หน่วยนั้นคูณด้วยตัวถ่วงซึ่งมีค่าต่างๆ ตามเอกสาร แนบท้ายสัญญาผนวก ๒", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ๖ การให้บริการ", null, "32", true));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้รับจ้างตกลงว่าการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ให้รวมถึงการบำรุงรักษาเพื่อป้องกันความชำรุดเสียหายของคอมพิวเตอร์ตลอดระยะเวลาตามสัญญานี้ โดยจะทำการซ่อมแซมแก้ไขและเปลี่ยนสิ่งที่จำเป็นทุกประการ เพื่อให้คอมพิวเตอร์อยู่ในสภาพใช้งานได้ดีตามปกติโดยไม่คิดค่าใช้จ่ายใดๆ เพิ่มเติมนอกเหนือจากค่าจ้างตามข้อ ๔ แห่งสัญญานี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้รับจ้างจะต้องจัดให้ช่างผู้มีความรู้ความชำนาญและฝีมือดีมาตรวจสอบบำรุงรักษาคอมพิวเตอร์ อย่างน้อยเดือนละ "+result.ServiceFixPerMonths+" ครั้ง ในกรณีคอมพิวเตอร์ขัดข้องใช้การไม่ได้ตามปกติ" +
                "ผู้รับจ้างจะต้องจัดการซ่อมแซมแก้ไขให้อยู่ในสภาพใช้การได้ดีดังเดิม โดยต้องเริ่มจัดการซ่อมแซมแก้ไขภายใน "+result.ServiceFixStartIn+"("+result.ServiceFixStartUnit+") วัน/ชั่วโมง นับตั้งแต่เวลาที่ได้รับแจ้งจากผู้ว่าจ้างหรือผู้ที่ได้รับมอบหมายจากผู้ว่าจ้างโดยจะแจ้งให้ผู้รับจ้างหรือผู้ที่ได้รับมอบหมายจากผู้รับจ้างทราบทางวาจา ทางโทรสาร หรือทางไปรษณีย์อิเล็กทรอนิกส์(e-mail) " +
                "หรือทางโทรศัพท์ ไม่ว่าวิธีใดวิธีหนึ่งให้ถือเป็นการแจ้งโดยชอบตามสัญญานี้แล้ว และผู้รับจ้างจะต้องซ่อมแซมแก้ไข หรือเปลี่ยนสิ่งที่จำเป็นให้เสร็จเรียบร้อยภายใน "+result.ServiceTimeIn+"("+result.ServiceTimeUnit+")  วัน/ชั่วโมง นับแต่เวลาที่ได้รับแจ้งจากผู้ว่าจ้างดังกล่าว", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้รับจ้างไม่เข้ามาซ่อมแซมแก้ไขภายในเวลาที่กำหนด หรือไม่สามารถดำเนินการซ่อมแซมแก้ไขหรือไม่สามารถจัดหาอุปกรณ์ใหม่ที่มีคุณสมบัติทัดเทียมกันหรือดีกว่ามาเปลี่ยนให้ใช้งานได้" +
                " ภายในเวลา ที่กำหนดไว้ ผู้รับจ้างยินยอมให้คิดค่าปรับเป็นรายชั่วโมง (เศษของชั่วโมงให้นับเป็น ๑ (หนึ่ง) ชั่วโมง) ในอัตราร้อยละ "+result.ServicePenaltyPercent+" ของค่าจ้างบำรุงรักษา (รายงวด) ตามสัญญา " +
                "นับจากเวลาที่ครบกำหนดจนถึงเวลาที่ผู้รับจ้างได้เริ่มการซ่อมแซมแก้ไข หรือจนถึงเวลาที่ผู้รับจ้างดำเนินการซ่อมแซมแก้ไขแล้วเสร็จแล้วแต่กรณี ทั้งนี้ หากผู้รับจ้างไม่ดำเนินการดังกล่าว ผู้ว่าจ้างมีสิทธิจ้างบุคคลภายนอกทำการซ่อมแซมแก้ไข " +
                "โดยผู้รับจ้างจะต้องออกค่าใช้จ่ายในการจ้างบุคคลภายนอกซ่อมแซมแก้ไขแทนผู้ว่าจ้างทั้งสิ้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("การจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ตามสัญญานี้ ไม่รวมถึงการเปลี่ยนแปลงลักษณะเฉพาะของคอมพิวเตอร์หรือส่วนประกอบที่ติดตั้งเพิ่มเติมภายหลังที่สัญญานี้มีผลบังคับและความเสียหายของคอมพิวเตอร์ซึ่งเกิดจากเหตุสุดวิสัยหรือเกิดจากความผิดของผู้ว่าจ้าง ", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ๗ ความรับผิดของผู้รับจ้าง", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้รับจ้างจะต้องรับผิดต่อผู้ว่าจ้างในกรณีที่ผู้รับจ้าง ผู้แทน ช่าง หรือลูกจ้างของผู้รับจ้างจงใจหรือประมาทเลินเล่อ หรือไม่มีความรู้ความชำนาญพอ " +
                "กระทำหรืองดเว้นการกระทำใดๆ เป็นเหตุให้คอมพิวเตอร์ของผู้ว่าจ้างเสียหายหรือไม่อยู่ในสภาพที่ใช้การได้ดี โดยไม่อาจแก้ไขได้ " +
                "โดยผู้รับจ้างจะต้องจัดหาคอมพิวเตอร์ ที่มีคุณภาพ ประสิทธิภาพ และความสามารถในการใช้งานไม่ต่ำกว่าของเดิมชดใช้แทน " +
                "หรือชดใช้ราคาคอมพิวเตอร์ในขณะที่เกิดความเสียหายในกรณีที่ไม่อาจจัดหาคอมพิวเตอร์ดังกล่าวชดใช้แทนได้ " +
                "ให้แก่ผู้ว่าจ้างภายในเวลาที่ ผู้ว่าจ้างกำหนด", null, "32"));

                string strContPenalty = CommonDAO.NumberToThaiText(result.ContPenaltyPerDays ?? 0);

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("นับตั้งแต่เวลาที่ผู้ว่าจ้างบอกกล่าวเป็นหนังสือให้ผู้รับจ้างจัดหาคอมพิวเตอร์มาชดใช้ให้แทน หรือชดใช้ราคาคอมพิวเตอร์ตามวรรคหนึ่ง " +
                "ผู้รับจ้างยินยอมให้ผู้ว่าจ้างปรับเป็นรายวันในอัตราร้อยละ "+result.ContPenaltyPercent+" ของค่าจ้างตามสัญญานี้ ซึ่งคิดเป็นเงิน "+ result.ContPenaltyPerDays + " บาท ("+ strContPenalty + ") ต่อวัน " +
                "จนกว่าผู้ว่าจ้างบอกเลิกสัญญาตามสัญญาข้อ๑๐ และหากผู้ว่าจ้างต้องใช้คอมพิวเตอร์ที่อื่นประมวลผลผู้รับจ้างยินยอมชดใช้ค่าใช้จ่ายเพื่อการดังกล่าวทั้งสิ้นแทนผู้ว่าจ้างอีกด้วย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("นอกจากนี้ ผู้รับจ้างจะต้องรับผิดชอบต่ออุบัติเหตุ ความเสียหาย หรือภยันตรายใดๆอันเกิดจากการปฏิบัติงานของผู้รับจ้าง" +
                "และต้องรับผิดต่อความเสียหายจากการกระทำหรืองดเว้นกระทำโดยผิดกฎหมายหรือโดยจงใจหรือประมาทเลินเล่อของผู้แทน ช่าง หรือลูกจ้างของผู้รับจ้างอีกด้วย", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ๘ หลักประกันการปฏิบัติตามสัญญา ", null, "32", true));
                string strGuaranteeAmount = CommonDAO.NumberToThaiText(result.GuaranteeAmount ?? 0);
               
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในขณะทำสัญญานี้ ผู้รับจ้างได้นำหลักประกันเป็น "+result.GuaranteeType+" เป็นจำนวนเงิน "+result.GuaranteeAmount +" บาท ("+ strGuaranteeAmount + ") ซึ่งเท่ากับร้อยละ "+result.GuaranteePercent+" ของราคา", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๑๑) กรณีผู้รับจ้างใช้หนังสือค้ำประกันมาเป็นหลักประกันการปฏิบัติตามสัญญา หนังสือค้ำประกันดังกล่าวจะต้องออกโดยธนาคารที่ประกอบกิจการในประเทศไทย หรือโดยบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบตามแบบที่คณะกรรมการนโยบายการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐกำหนด หรืออาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้ และจะต้องมีอายุการค้ำประกันตลอดไปจนกว่าผู้รับจ้างพ้นข้อผูกพันตามสัญญานี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หลักประกันที่ผู้รับจ้างนำมามอบให้ตามวรรคหนึ่งจะต้องมีอายุครอบคลุมความรับผิดทั้งปวงของผู้รับจ้างตลอดอายุสัญญาถ้าหลักประกันที่ผู้รับจ้างนำมามอบให้ดังกล่าวลดลงหรือเสื่อมค่าลง " +
                    "หรือมีอายุ ไม่ครอบคลุมถึงความรับผิดของผู้รับจ้างตลอดอายุสัญญา ไม่ว่าด้วยเหตุใดๆ ก็ตาม ผู้รับจ้างจะต้องหาหลักประกันมาเปลี่ยนให้ใหม่หรือหลักประกันเพิ่มเติมให้มีจำนวนครบถ้วนตามวรรคหนึ่งมามอบให้แก่ผู้ว่าจ้างภายใน "+result.EnforcementOfFineDays+" วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้าง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หลักประกันที่ผู้รับจ้างนำมามอบไว้ตามข้อนี้ ผู้ว่าจ้างจะคืนให้แก่ผู้รับจ้างโดยไม่มีดอกเบี้ยเมื่อผู้รับจ้างพ้นจากข้อผูกพันและความรับผิดทั้งปวงตามสัญญานี้แล้ว", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๙ การจ้างช่วง", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้รับจ้างจะต้องไม่เอางานทั้งหมดหรือแต่บางส่วนแห่งสัญญานี้ไปจ้างช่วงอีกทอดหนึ่ง " +
                "เว้นแต่การจ้างช่วงงานแต่บางส่วนที่ได้รับอนุญาตเป็นหนังสือจากผู้ว่าจ้างแล้ว การที่ผู้ว่าจ้างได้อนุญาตให้จ้างช่วงงาน แต่บางส่วนดังกล่าวนั้น ไม่เป็นเหตุให้ผู้รับจ้างหลุดพ้นจากความรับผิดหรือพันธะหน้าที่ตามสัญญานี้ และผู้รับจ้าง " +
                "จะยังคงต้องรับผิดในความผิดและความประมาทเลินเล่อของผู้รับจ้างช่วง หรือของตัวแทนหรือลูกจ้างของผู้รับจ้างช่วงนั้นทุกประการ ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("กรณีผู้รับจ้างไปจ้างช่วงงานแต่บางส่วนโดยฝ่าฝืนความในวรรคหนึ่ง ผู้รับจ้างต้องชำระค่าปรับให้แก่ผู้ว่าจ้างเป็นจำนวนเงินใน" +
                "อัตราร้อยละ.......(๑๒)…...…(...................) ของวงเงินของงานที่จ้างช่วงตามสัญญา ทั้งนี้ ไม่ตัดสิทธิผู้ว่าจ้างในการบอกเลิกสัญญา", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๐ การบอกเลิกสัญญา", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หากผู้ว่าจ้างเห็นว่าผู้รับจ้างไม่อาจปฏิบัติตามสัญญาได้ หรือผู้รับจ้างผิดสัญญาข้อใดข้อหนึ่ง" +
                "หรือตกเป็นผู้ล้มละลาย ผู้ว่าจ้างมีสิทธิบอกเลิกสัญญาได้ และมีสิทธิจ้างผู้รับจ้างรายใหม่เข้าทำงานของผู้รับจ้างให้ลุล่วงไป" +
                " การใช้สิทธิบอกเลิกสัญญานั้นไม่กระทบสิทธิของผู้ว่าจ้างที่จะเรียกร้องค่าเสียหายและค่าใช้จ่ายใดๆ (ถ้ามี) จากผู้รับจ้าง ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้ว่าจ้างใช้สิทธิบอกเลิกสัญญา ผู้ว่าจ้างมีสิทธิริบหรือบังคับจากหลักประกัน" +
                "การปฏิบัติตามสัญญาตามข้อ๘ เป็นจำนวนเงินทั้งหมดหรือแต่บางส่วน ตามแต่จะเห็นสมควรได้ทันที นอกจากนั้น ผู้รับจ้างจะต้องรับผิดชอบในค่าเสียหายซึ่งเป็นจำนวนเกินกว่าหลักประกันการปฏิบัติตามสัญญา " +
                "และค่าเสียหายต่างๆ ที่เกิดขึ้น รวมทั้งค่าใช้จ่ายที่เพิ่มขึ้นในการทำงานนั้นต่อให้แล้วเสร็จตามสัญญาซึ่งผู้ว่าจ้างจะหักเอาจากจำนวนเงินใดๆ ที่จะจ่ายให้แก่ผู้รับจ้างก็ได้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("การที่ผู้ว่าจ้างไม่ใช้สิทธิเลิกสัญญาดังกล่าวตามวรรคหนึ่งไม่เป็นเหตุให้ผู้รับจ้างพ้นจากความรับผิดตามสัญญา", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๑ การบังคับค่าปรับ ค่าเสียหาย และค่าใช้จ่าย", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้รับจ้างไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่งด้วยเหตุใดๆ ก็ตาม จนเป็นเหตุให้เกิดค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายแก่ผู้ว่าจ้าง ผู้รับจ้างต้องชดใช้ค่าปรับ ค่าเสียหาย" +
                    " หรือค่าใช้จ่ายดังกล่าวให้แก่ผู้ว่าจ้างโดยสิ้นเชิงภายในกำหนด....................(......................) วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้าง หากผู้รับจ้างไม่ชดใช้ให้ถูกต้องครบถ้วนภายในระยะเวลาดังกล่าวให้ผู้ว่าจ้างมีสิทธิที่จะหักเอาจากจำนวนเงินค่าจ้าง ที่ต้องชำระหรือบังคับจากหลักประกันการปฏิบัติตามสัญญาได้ทันที", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หากค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายที่บังคับจากจำนวนเงินค่าจ้างที่ต้องชำระ  หรือจากหลักประกันการปฏิบัติตามสัญญาแล้วยังไม่เพียงพอ ผู้รับจ้างยินยอมชำระส่วนที่เหลือที่ยังขาดอยู่จนครบถ้วน ตามจำนวนค่าปรับ ค่าเสียหาย " +
                    "หรือค่าใช้จ่ายนั้น ภายในกำหนด.............(.................) วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้าง", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ ๑๒การงดหรือลดค่าปรับ หรือการขยายเวลาในการปฏิบัติตามสัญญา", null, "32", true));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่มีเหตุเกิดจากความผิดหรือความบกพร่องของฝ่ายผู้ว่าจ้าง หรือเหตุสุดวิสัย หรือเกิดจากพฤติการณ์อันหนึ่งอันใดที่ผู้รับจ้างไม่ต้องรับผิดตามกฎหมาย หรือเหตุอื่นตามที่กำหนดในกฎกระทรวง " +
                "ซึ่งออกตามความในกฎหมายว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ ทำให้ผู้รับจ้างไม่สามารถปฏิบัติตามเงื่อนไขและกำหนดเวลาในข้อ ๕ ข้อ ๖ หรือข้อ ๗ ได้ ผู้รับจ้างจะต้องแจ้งเหตุหรือพฤติการณ์ดังกล่าว " +
                "พร้อมหลักฐานเป็นหนังสือให้ผู้ว่าจ้างทราบ เพื่อของดหรือลดค่าปรับ หรือขยายเวลาทำการตามสัญญาภายใน ๑๕ (สิบห้า) วันนับถัดจากวันที่เหตุนั้นสิ้นสุดลง หรือตามที่กำหนดในกฎกระทรวงดังกล่าว แล้วแต่กรณี", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ถ้าผู้รับจ้างไม่ปฏิบัติให้เป็นไปตามความในวรรคหนึ่ง ให้ถือว่าผู้รับจ้างได้สละสิทธิเรียกร้องในการที่จะของดหรือลดค่าปรับ" +
                "หรือขยายเวลาทำการตามสัญญาโดยไม่มีเงื่อนไขใดๆทั้งสิ้น เว้นแต่กรณีเหตุเกิดจากความผิดหรือความบกพร่องของฝ่ายผู้ว่าจ้างซึ่งมีหลักฐานชัดแจ้งหรือผู้ว่าจ้างทราบดีอยู่แล้วตั้งแต่ต้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("การงดหรือลดค่าปรับ หรือขยายกำหนดเวลาทำการตามสัญญาตามวรรคหนึ่ง อยู่ในดุลพินิจของผู้ว่าจ้างที่จะพิจารณาตามที่เห็นสมควร", null, "32"));

                body.AppendChild(WordServiceSetting.EmptyParagraph());

                body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ........................................................................ผู้ว่าจ้าง"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(................................................................................)"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ........................................................................ผู้ว่าจ้าง"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(................................................................................)"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ......................................................................พยาน"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(...............................................................................)"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ......................................................................พยาน"));
                body.AppendChild(WordServiceSetting.CenteredParagraph("(...............................................................................)"));

                // next page
                body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("วิธีปฏิบัติเกี่ยวกับสัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์", "000000", "36"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(1) ให้ระบุเลขที่สัญญาในปีงบประมาณหนึ่งๆ ตามลำดับ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(2) ให้ระบุชื่อของหน่วยงานของรัฐที่เป็นนิติบุคคล เช่น กรม ก. หรือรัฐวิสาหกิจ ข. เป็นต้น", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(3) ให้ระบุชื่อและตำแหน่งของหัวหน้าหน่วยงานของรัฐที่เป็นนิติบุคคลนั้น หรือผู้ที่ได้รับมอบอำนาจ เช่น นาย ก. อธิบดีกรม………...… หรือ นาย ข. ผู้ได้รับมอบอำนาจจากอธิบดีกรม………......………..", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(4) ให้ระบุชื่อผู้รับจ้าง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ก. กรณีนิติบุคคล เช่น ห้างหุ้นส่วนสามัญจดทะเบียน ห้างหุ้นส่วนจำกัด บริษัทจำกัด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ข. กรณีบุคคลธรรมดา ให้ระบุชื่อและที่อยู่", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(5) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(6) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(7) ให้กำหนดในอัตราระหว่างร้อยละ ๐.๐๒๕ – ๐.๐๓๕ ของราคาตามสัญญาต่อชั่วโมง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(8) อัตราค่าปรับตามสัญญาข้อ ๗ ให้กำหนดเป็นรายวันในอัตราระหว่างร้อยละ ๐.๐๑-๐.๑๐ " +
                  "ตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ ข้อ ๑๖๒ " +
                  "ส่วนกรณีจะปรับร้อยละเท่าใด ให้อยู่ในดุลพินิจของหน่วยงานของรัฐผู้ว่าจ้างที่จะพิจารณา" +
                  " โดยคำนึงถึงราคาและลักษณะของพัสดุที่จ้าง ซึ่งอาจมีผลกระทบต่อการที่ผู้รับจ้างจะหลีกเลี่ยงไม่ปฏิบัติตามสัญญา " +
                  "แต่ทั้งนี้การที่จะกำหนดค่าปรับเป็นร้อยละเท่าใดจะต้องกำหนดไว้ในเอกสารเชิญชวนด้วย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(9)“หลักประกัน” หมายถึง หลักประกันที่ผู้รับจ้างนำมามอบไว้แก่หน่วยงานของรัฐ เมื่อลงนามในสัญญา เพื่อเป็นการประกันความเสียหายที่อาจจะเกิดขึ้นจากการปฏิบัติตามสัญญา ดังนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๑)เงินสด ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๒)เช็คหรือดราฟท์ที่ธนาคารเซ็นสั่งจ่าย ซึ่งเป็นเช็คหรือดราฟท์ลงวันที่ที่ใช้เช็คหรือดราฟท์นั้นชำระต่อเจ้าหน้าที่ หรือก่อนวันนั้นไม่เกิน ๓ วันทำการ ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๓)หนังสือคํ้าประกันของธนาคารภายในประเทศตามตัวอย่างที่คณะกรรมการนโยบายกําหนด โดยอาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๔)หนังสือค้ำประกันของบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทยตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบ โดยอนุโลมให้ใช้ตามตัวอย่างหนังสือค้ำประกันของธนาคารที่คณะกรรมการนโยบายกําหนด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๕)พันธบัตรรัฐบาลไท", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(10) ให้กำหนดจำนวนเงินหลักประกันการปฏิบัติตามสัญญาตามระเบียบการะทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ ข้อ ๑๖๘", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(11) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(12) อัตราค่าปรับตามสัญญาข้อ ๙ กรณีผู้รับจ้างไปจ้างช่วงบางส่วนโดยไม่ได้รับอนุญาตจากผู้ว่าจ้างต้องกำหนดค่าปรับเป็นจำนวนเงินไม่น้อยกว่าร้อยละสิบของวงเงินของงานที่จ้างช่วงตามสัญญา", null, "32"));






                body.AppendChild(WordServiceSetting.EmptyParagraph());




                WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);

            }
            stream.Position = 0;
            return stream.ToArray();
        }

    
    }

    public async Task<byte[]> OnGetWordContact_MaintenanceComputer_ToPDF(string id,string typeContact)
    {
        try
        {
            var result = await _econtractReportSMCDAO.GetSMCAsync(id);
            if (result == null)
            {
                throw new Exception("SMC data not found.");
            }
            else
            {
          
                if (result == null)
                    throw new Exception("SMC data not found.");

                var installmentList = await _econtractReportSMCDAO.GetSMCInstallmentAsync(id);

                var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");

                var listDocAtt = await _eContractDAO.GetRelatedDocumentsAsync(id, "SMC31060");
                var htmlDocAtt = listDocAtt != null
                    ? string.Join("", listDocAtt.Select(docItem =>
                        $"<p class='tab3 t-16'>{docItem.DocumentTitle} จำนวน {docItem.PageAmount} หน้า</div>"))
                    : "";

                string datestring = CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now);
                string strServiceStartDate = CommonDAO.ToThaiDateStringCovert(result.ServiceStartDate ?? DateTime.Now);
                string strServiceEndDate = CommonDAO.ToThaiDateStringCovert(result.ServiceStartDate ?? DateTime.Now);
                string strServiceFee = CommonDAO.NumberToThaiText(result.ServiceFee ?? 0);
                string strVatAmount = CommonDAO.NumberToThaiText(result.VatAmount ?? 0);
                string strPenaltyPerHours = CommonDAO.NumberToThaiText(result.PenaltyPerHours ?? 0);
                string strContPenalty = CommonDAO.NumberToThaiText(result.ContPenaltyPerDays ?? 0);
                string strGuaranteeAmount = CommonDAO.NumberToThaiText(result.GuaranteeAmount ?? 0);

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


                // Build HTML for installment list
                string htmlInstallments = "";
                if (installmentList != null && installmentList.ToList().Count > 0)
                {
                    foreach (var item in installmentList)
                    {
                        htmlInstallments += $"<p class='tab3 t-16'>งวดที่ {item.PayRound} เป็นเงิน  {item.TotalAmount} บาท จะจ่ายเมื่อผู้รับจ้างได้ดำเนินการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์เป็นเวลา {item.RepairMonth} เดือน และผู้ว่าจ้างได้ตรวจรับมอบงานตามสัญญาแล้ว</div>";
                    }
                }

                var htmlBody = $@"
<div class='contract'>
    <div class='text-center t-22' style='font-weight:bold;'>แบบสัญญา</div>
    <div class='text-center t-18' style='font-weight:bold;'>สัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์</div>
</br>
    <div class='tab2 t-16 text-right'>สัญญาเลขที่…{result.Contract_Number ?? "xxxxx"}.</div>
    <p class='tab2 t-16'>
        สัญญาฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.)
        ตำบล/แขวง
</br>ทุ่งสองห้อง อำเภอ/เขต หลักสี่
        จังหวัด กรุงเทพ เมื่อ {datestring}
        ระหว่าง {result.Contract_Organization}
        โดย {result.SignatoryName}
        ซึ่งต่อไปในสัญญานี้เรียกว่า “ผู้ซื้อ” ฝ่ายหนึ่ง กับ…{result.ContractorName}
    </p>
    {(result.ContractorType == "นิติบุคคล"
                    ? $@"<p class='tab2 t-16'>
            ซึ่งจดทะเบียนเป็นนิติบุคคล ณ {result.ContractorName} มี
            สำนักงานใหญ่อยู่เลขที่ {result.ContractorAddressNo} ถนน {result.ContractorStreet} ตำบล/แขวง {result.ContractorSubDistrict}
            อำเภอ/เขต {result.ContractorDistrict} จังหวัด {result.ContractorProvince}
            โดย {result.ContractorSignatoryName}
            มีอำนาจลงนามผูกพันนิติบุคคลปรากฏตามหนังสือรับรองของสำนักงานทะเบียนหุ้นส่วนบริษัท ……………
            ลงวันที่ {datestring} (5)(และหนังสือมอบอำนาจลง {datestring}) แนบท้ายสัญญานี้
        </p>"
                    : $@"<p class='tab2 t-16'>
            (6)(ในกรณีที่ผู้รับจ้างเป็นบุคคลธรรมดาให้ใช้ข้อความว่า กับ {result.ContractorName}
            อยู่บ้านเลขที่ {result.ContractorAddressNo} ถนน {result.ContractorStreet} ตำบล/แขวง {result.ContractorSubDistrict}
            อำเภอ/เขต {result.ContractorDistrict} จังหวัด {result.ContractorProvince}
            ผู้ถือบัตรประจำตัวประชาชนเลขที่ {result.CitizenId} ดังปรากฏตามสำเนาบัตรประจำตัวประชาชนแนบท้ายสัญญานี้) ซึ่งต่อไปในสัญญานี้เรียกว่า “ผู้ซื้อ” อีกฝ่ายหนึ่ง กับ…{result.ContractorName}
        </p>")}
    <p class='tab2 t-16'>คู่สัญญาได้ตกลงกันมีข้อความดังต่อไปนี้</p>
    <p class='tab2 t-16'><b>ข้อ ๑ ขอบเขตของงาน</b></p>
   <p class='tab3 t-16'>ผู้ว่าจ้างตกลงจ้างและผู้รับจ้างตกลงรับจ้างให้บริการบำรุงรักษาและซ่อมแซมแก้ไข
</br>คอมพิวเตอร์ อุปกรณ์การประมวลผล และระบบคอมพิวเตอร์ ตามเอกสารแนบท้ายสัญญาผนวก ๑ ซึ่งต่อไปในสัญญานี้เรียกว่า “คอมพิวเตอร์” ซึ่งติดตั้งอยู่ ณ {result.Contract_Sign_Address}</p>
   <p class='tab3 t-16'>ผู้รับจ้างจะจัดหาวัสดุสิ่งของชนิดที่ดีได้มาตรฐาน ใช้เครื่องมือดี และช่างผู้ชำนาญและฝีมือดี
</br>เพื่อใช้ในงานจ้างที่จำเป็นสำหรับการปฏิบัติงานตามสัญญา</p>
    <p class='tab2 t-16'><b>ข้อ ๒ เอกสารอันเป็นส่วนหนึ่งของสัญญา</b></p>
   <p class='tab3 t-16'>เอกสารแนบท้ายสัญญาดังต่อไปนี้ให้ถือว่าเป็นส่วนหนึ่งของสัญญานี้</p>
{htmlDocAtt}
   <p class='tab3 t-16'>และอัตราค่าบำรุงรักษา</p>
   <p class='tab3 t-16'>ความใดในเอกสารแนบท้ายสัญญาที่ขัดหรือแย้งกับข้อความในสัญญานี้ ให้ใช้ข้อความใน
</br>สัญญานี้บังคับ และในกรณีที่เอกสารแนบท้ายสัญญาขัดแย้งกันเอง ผู้รับจ้างจะต้องปฏิบัติตามคำวินิจฉัยของ ผู้ว่าจ้าง คำวินิจฉัยของผู้ว่าจ้างให้ถือเป็นที่สุด และผู้รับจ้างไม่มีสิทธิเรียกร้องค่าจ้าง ค่าเสียหาย หรือค่าใช้
</br>จ่ายใดๆ เพิ่มเติมจากผู้ว่าจ้างทั้งสิ้น</p>
    <p class='tab2 t-16'><b>ข้อ ๓ ระยะเวลาให้บริการ</b></p>
   <p class='tab3 t-16'>ผู้รับจ้างตกลงให้บริการตามสัญญานี้ ตั้งแต่วันที่ {strServiceStartDate} ถึง {strServiceEndDate} รวมเป็นเวลาทั้งสิ้น {result.ServiceTotalYears} ปี {result.ServiceTotalMonths} เดือน {result.ServiceTotalDays} วัน</p>
    <p class='tab2 t-16'><b>ข้อ ๔ ค่าจ้างและการจ่ายเงิน</b></p>
   <p class='tab3 t-16'>ผู้ว่าจ้างตกลงชำระค่าจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์เป็นเงินทั้งสิ้น {result.ServiceFee} บาท ({strServiceFee}) ซึ่งได้รวมภาษีมูลค่าเพิ่มจำนวน {result.VatAmount} บาท ({strVatAmount}) ตลอดจนค่าแรงงานค่าสิ่งของตลอดอายุสัญญา ภาษีอากรอื่น และค่าใช้จ่ายทั้งปวงไว้ด้วยแล้ว โดยผู้ว่าจ้างจะแบ่งจ่ายให้แก่ผู้รับจ้างเป็นงวดๆ รวม {result.PaymentInstallment} งวด ดังนี้</p>
    {htmlInstallments}
<!-- ข้อ 5 การรับประกันผลงาน -->
<p class='tab2 t-16'><b>ข้อ ๕ การรับประกันผลงาน</b></p>
<p class='tab3 t-16'>
    ผู้รับจ้างตกลงบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ตามสัญญานี้ให้อยู่ในสภาพใช้งานได้ดีอยู่เสมอ โดยให้มีเวลาคอมพิวเตอร์ขัดข้องรวมตามเกณฑ์การคำนวณเวลาขัดข้อง ไม่เกินเดือนละ {result.MaximumDownTimeHours} ชั่วโมง 
</br>หรือร้อยละ {result.MaximumDownPercents} ของเวลาใช้งานทั้งหมดของคอมพิวเตอร์ของเดือนนั้น แล้วแต่ตัวเลขใดจะมากกว่ากัน 
</br>มิฉะนั้นผู้รับจ้างต้องยอมให้ผู้ว่าจ้างคิดค่าปรับเป็นรายชั่วโมง ในอัตราชั่วโมงละ {result.PenaltyPerHours} บาท ({strPenaltyPerHours}) ในช่วงเวลาที่ไม่สามารถใช้คอมพิวเตอร์ได้ในส่วนที่เกินกว่ากำหนดเวลา
</br>ขัดข้องข้างต้น
</p>
<p class='tab3 t-16'>เกณฑ์การคำนวณเวลาขัดข้องของคอมพิวเตอร์ตามวรรคหนึ่งให้เป็นไปดังนี้</p>
<p class='tab3 t-16'>- กรณีที่คอมพิวเตอร์เกิดขัดข้องพร้อมกันหลายหน่วย ให้นับเวลาขัดข้องของหน่วยที่มีตัวถ่วงมากที่สุดเพียงหน่วยเดียว</p>
<p class='tab3 t-16'>- กรณีความเสียหายอันสืบเนื่องมาจากความขัดข้องของคอมพิวเตอร์แตกต่างกัน เวลาที่ใช้ในการคำนวณค่าปรับจะเท่ากับเวลาขัดข้องของคอมพิวเตอร์หน่วยนั้นคูณด้วยตัวถ่วงซึ่งมีค่าต่างๆ ตามเอกสาร
</br>แนบท้ายสัญญาผนวก ๒</p>

<!-- ข้อ 6 การให้บริการ -->
<p class='tab2 t-16'><b>ข้อ ๖ การให้บริการ</b></p>
<p class='tab3 t-16'>
    ผู้รับจ้างตกลงว่าการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ให้รวมถึงการบำรุงรักษาเพื่อ
</br>ป้องกันความชำรุดเสียหายของคอมพิวเตอร์ตลอดระยะเวลาตามสัญญานี้ โดยจะทำการซ่อมแซมแก้ไขและ
</br>เปลี่ยนสิ่งที่จำเป็นทุกประการ เพื่อให้คอมพิวเตอร์อยู่ในสภาพใช้งานได้ดีตามปกติโดยไม่คิดค่าใช้จ่ายใดๆ
</br>เพิ่มเติมนอกเหนือจากค่าจ้างตามข้อ ๔ แห่งสัญญานี้
</p>
<p class='tab3 t-16'>
    ผู้รับจ้างจะต้องจัดให้ช่างผู้มีความรู้ความชำนาญและฝีมือดีมาตรวจสอบบำรุงรักษา
</br>คอมพิวเตอร์ อย่างน้อยเดือนละ {result.ServiceFixPerMonths} ครั้ง ในกรณีคอมพิวเตอร์ขัดข้องใช้การไม่ได้ตามปกติ ผู้รับจ้างจะต้อง
</br>จัดการซ่อมแซมแก้ไขให้อยู่ในสภาพใช้การได้ดีดังเดิม โดยต้องเริ่มจัดการซ่อมแซมแก้ไขภายใน {result.ServiceFixStartIn}({result.ServiceFixStartUnit}) วัน/
</br>ชั่วโมงนับตั้งแต่เวลาที่ได้รับแจ้งจากผู้ว่าจ้างหรือผู้ที่ได้รับมอบหมายจากผู้ว่าจ้างโดยจะแจ้งให้ผู้รับจ้างหรือผู้ที่
</br>ได้รับมอบหมายจากผู้รับจ้างทราบทางวาจา ทางโทรสาร หรือทางไปรษณีย์อิเล็กทรอนิกส์(e-mail) หรือทางโทรศัพท์ ไม่ว่าวิธีใดวิธีหนึ่งให้ถือเป็นการแจ้งโดยชอบตามสัญญานี้แล้ว และผู้รับจ้างจะต้องซ่อมแซมแก้ไข 
</br>หรือเปลี่ยนสิ่งที่จำเป็นให้เสร็จเรียบร้อยภายใน {result.ServiceTimeIn}({result.ServiceTimeUnit}) วัน/ชั่วโมง นับแต่เวลาที่ได้รับแจ้งจากผู้ว่าจ้าง
</br>ดังกล่าว
</p>
<p class='tab3 t-16'>
    ในกรณีที่ผู้รับจ้างไม่เข้ามาซ่อมแซมแก้ไขภายในเวลาที่กำหนด หรือไม่สามารถดำเนินการซ่อมแซมแก้ไขหรือไม่สามารถจัดหาอุปกรณ์ใหม่ที่มีคุณสมบัติทัดเทียมกันหรือดีกว่ามาเปลี่ยนให้ใช้งานได้ ภายใน
</br>เวลาที่กำหนดไว้ ผู้รับจ้างยินยอมให้คิดค่าปรับเป็นรายชั่วโมง (เศษของชั่วโมงให้นับเป็น ๑ (หนึ่ง) ชั่วโมง) 
</br>ในอัตราร้อยละ {result.ServicePenaltyPercent} ของค่าจ้างบำรุงรักษา (รายงวด) ตามสัญญา นับจากเวลาที่ครบกำหนดจนถึงเวลาที่ผู้
</br>รับจ้างได้เริ่มการซ่อมแซมแก้ไข หรือจนถึงเวลาที่ผู้รับจ้างดำเนินการซ่อมแซมแก้ไขแล้วเสร็จแล้วแต่กรณี 
</br>ทั้งนี้ หากผู้รับจ้างไม่ดำเนินการดังกล่าว ผู้ว่าจ้างมีสิทธิจ้างบุคคลภายนอกทำการซ่อมแซมแก้ไข 
</br>โดยผู้รับจ้างจะต้องออกค่าใช้จ่ายในการจ้างบุคคลภายนอกซ่อมแซมแก้ไขแทนผู้ว่าจ้างทั้งสิ้น
</p>
<p class='tab3 t-16'>
    การจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ตามสัญญานี้ ไม่รวมถึงการเปลี่ยน
</br>แปลงลักษณะเฉพาะของคอมพิวเตอร์หรือส่วนประกอบที่ติดตั้งเพิ่มเติมภายหลังที่สัญญานี้มีผลบังคับ
</br>และความเสียหายของคอมพิวเตอร์ซึ่งเกิดจากเหตุสุดวิสัยหรือเกิดจากความผิดของผู้ว่าจ้าง
</p>

<!-- ข้อ 7 ความรับผิดของผู้รับจ้าง -->
<p class='tab2 t-16'><b>ข้อ ๗ ความรับผิดของผู้รับจ้าง</b></p>
<p class='tab3 t-16'>
    ผู้รับจ้างจะต้องรับผิดต่อผู้ว่าจ้างในกรณีที่ผู้รับจ้าง ผู้แทน ช่าง หรือลูกจ้างของผู้รับจ้างจงใจ
</br>หรือประมาทเลินเล่อ หรือไม่มีความรู้ความชำนาญพอ กระทำหรืองดเว้นการกระทำใดๆ เป็นเหตุให้
</br>คอมพิวเตอร์ของผู้ว่าจ้างเสียหายหรือไม่อยู่ในสภาพที่ใช้การได้ดี โดยไม่อาจแก้ไขได้ โดยผู้รับจ้างจะต้อง
</br>จัดหาคอมพิวเตอร์ ที่มีคุณภาพ ประสิทธิภาพ และความสามารถในการใช้งานไม่ต่ำกว่าของเดิมชดใช้แทน 
</br>หรือชดใช้ราคาคอมพิวเตอร์ในขณะที่เกิดความเสียหายในกรณีที่ไม่อาจจัดหาคอมพิวเตอร์ดังกล่าวชดใช้แทนได้ 
</br>ให้แก่ผู้ว่าจ้างภายในเวลาที่ ผู้ว่าจ้างกำหนด
</p>
<p class='tab3 t-16'>
    นับตั้งแต่เวลาที่ผู้ว่าจ้างบอกกล่าวเป็นหนังสือให้ผู้รับจ้างจัดหาคอมพิวเตอร์มาชดใช้ให้แทน 
</br>หรือชดใช้ราคาคอมพิวเตอร์ตามวรรคหนึ่ง ผู้รับจ้างยินยอมให้ผู้ว่าจ้างปรับเป็นรายวันในอัตราร้อยละ {result.ContPenaltyPercent} 
</br>ของค่าจ้างตามสัญญานี้ซึ่งคิดเป็นเงิน {result.ContPenaltyPerDays} บาท ({strContPenalty}) ต่อวัน 
</br>จนกว่าผู้ว่าจ้างบอกเลิกสัญญาตามสัญญาข้อ ๑๐ และหากผู้ว่าจ้างต้องใช้คอมพิวเตอร์ที่อื่นประมวลผล
</br>ผู้รับจ้างยินยอมชดใช้ค่าใช้จ่ายเพื่อการดังกล่าวทั้งสิ้นแทนผู้ว่าจ้างอีกด้วย
</p>
<p class='tab3 t-16'>
    นอกจากนี้ ผู้รับจ้างจะต้องรับผิดชอบต่ออุบัติเหตุ ความเสียหาย หรือภยันตรายใดๆอันเกิด
</br>จากการปฏิบัติงานของผู้รับจ้าง และต้องรับผิดต่อความเสียหายจากการกระทำหรืองดเว้นกระทำโดยผิด
</br>กฎหมายหรือโดยจงใจหรือประมาทเลินเล่อของผู้แทน ช่าง หรือลูกจ้างของผู้รับจ้างอีกด้วย
</p>

<!-- ข้อ 8 หลักประกันการปฏิบัติตามสัญญา -->
<p class='tab2 t-16'><b>ข้อ ๘ หลักประกันการปฏิบัติตามสัญญา</b></p>
<p class='tab3 t-16'>
    ในขณะทำสัญญานี้ ผู้รับจ้างได้นำหลักประกันเป็น {result.GuaranteeType} เป็นจำนวนเงิน {result.GuaranteeAmount} บาท ({strGuaranteeAmount}) ซึ่งเท่ากับร้อยละ {result.GuaranteePercent} ของราคา
</p>
<p class='tab3 t-16'>
    (๑๑) กรณีผู้รับจ้างใช้หนังสือค้ำประกันมาเป็นหลักประกันการปฏิบัติตามสัญญา หนังสือค้ำ
</br>ประกันดังกล่าวจะต้องออกโดยธนาคารที่ประกอบกิจการในประเทศไทย หรือโดยบริษัทเงินทุนหรือบริษัท
</br>เงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตาม
</br>ประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบ
</br>ตามแบบที่คณะกรรมการนโยบายการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐกำหนด หรืออาจเป็นหนังสือ
</br>ค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้และจะต้องมีอายุการค้ำประกันตลอด
</br>ไปจนกว่าผู้รับจ้างพ้นข้อผูกพันตามสัญญานี้
</p>
<p class='tab3 t-16'>
    หลักประกันที่ผู้รับจ้างนำมามอบให้ตามวรรคหนึ่งจะต้องมีอายุครอบคลุมความรับผิดทั้งปวง
</br>ของผู้รับจ้างตลอดอายุสัญญา ถ้าหลักประกันที่ผู้รับจ้างนำมามอบให้ดังกล่าวลดลงหรือเสื่อมค่าลง 
</br>หรือมีอายุไม่ครอบคลุมถึงความรับผิดของผู้รับจ้างตลอดอายุสัญญา ไม่ว่าด้วยเหตุใดๆ ก็ตาม 
</br>ผู้รับจ้างจะต้องหาหลักประกันมาเปลี่ยนให้ใหม่หรือหลักประกันเพิ่มเติมให้มีจำนวนครบถ้วนตามวรรคหนึ่ง
</br>มามอบให้แก่ผู้ว่าจ้างภายใน {result.EnforcementOfFineDays} วัน นับถัดจากวันที่ได้รับแจ้ง
</br>เป็นหนังสือจากผู้ว่าจ้าง
</p>
<p class='tab3 t-16'>
    หลักประกันที่ผู้รับจ้างนำมามอบไว้ตามข้อนี้ ผู้ว่าจ้างจะคืนให้แก่ผู้รับจ้างโดยไม่มีดอกเบี้ยเมื่อผู้รับจ้างพ้นจากข้อผูกพันและความรับผิดทั้งปวงตามสัญญานี้แล้ว
</p>

<!-- ข้อ 9 การจ้างช่วง -->
<p class='tab2 t-16'><b>ข้อ ๙ การจ้างช่วง</b></p>
<p class='tab3 t-16'>
    ผู้รับจ้างจะต้องไม่เอางานทั้งหมดหรือแต่บางส่วนแห่งสัญญานี้ไปจ้างช่วงอีกทอดหนึ่ง เว้นแต่
</br>การจ้างช่วงงานแต่บางส่วนที่ได้รับอนุญาตเป็นหนังสือจากผู้ว่าจ้างแล้ว การที่ผู้ว่าจ้างได้อนุญาตให้จ้างช่วงงาน 
</br>แต่บางส่วนดังกล่าวนั้น ไม่เป็นเหตุให้ผู้รับจ้างหลุดพ้นจากความรับผิดหรือพันธะหน้าที่ตามสัญญานี้ 
</br>และผู้รับจ้างจะยังคงต้องรับผิดในความผิดและความประมาทเลินเล่อของผู้รับจ้างช่วง 
</br>หรือของตัวแทนหรือลูกจ้างของผู้รับจ้างช่วงนั้นทุกประการ
</p>
<p class='tab3 t-16'>
    กรณีผู้รับจ้างไปจ้างช่วงงานแต่บางส่วนโดยฝ่าฝืนความในวรรคหนึ่ง ผู้รับจ้างต้องชำระค่าปรับให้แก่ผู้ว่าจ้างเป็นจำนวนเงินในอัตราร้อยละ{result.SubcontractPenaltyPercent} ของวงเงินของงานที่จ้างช่วงตามสัญญา ทั้งนี้ ไม่ตัดสิทธิผู้ว่า
</br>จ้างในการบอกเลิกสัญญา
</p>

<!-- ข้อ 10 การบอกเลิกสัญญา -->
<p class='tab2 t-16'><b>ข้อ ๑๐ การบอกเลิกสัญญา</b></p>
<p class='tab3 t-16'>
    หากผู้ว่าจ้างเห็นว่าผู้รับจ้างไม่อาจปฏิบัติตามสัญญาได้ หรือผู้รับจ้างผิดสัญญาข้อใดข้อหนึ่ง 
</br>หรือตกเป็นผู้ล้มละลาย ผู้ว่าจ้างมีสิทธิบอกเลิกสัญญาได้ และมีสิทธิจ้างผู้รับจ้างรายใหม่เข้าทำงานของผู้รับจ้าง
</br>ให้ลุล่วงไป การใช้สิทธิบอกเลิกสัญญานั้นไม่กระทบสิทธิของผู้ว่าจ้างที่จะเรียกร้องค่าเสียหายและค่าใช้จ่ายใดๆ (ถ้ามี) จากผู้รับจ้าง
</p>
<p class='tab3 t-16'>
    ในกรณีที่ผู้ว่าจ้างใช้สิทธิบอกเลิกสัญญา ผู้ว่าจ้างมีสิทธิริบหรือบังคับจากหลักประกันการ
</br>ปฏิบัติตามสัญญาตามข้อ ๘ เป็นจำนวนเงินทั้งหมดหรือแต่บางส่วน ตามแต่จะเห็นสมควรได้ทันที นอกจากนั้น 
</br>ผู้รับจ้างจะต้องรับผิดชอบในค่าเสียหายซึ่งเป็นจำนวนเกินกว่าหลักประกันการปฏิบัติตามสัญญา และค่าเสีย
</br>หายต่างๆ ที่เกิดขึ้น รวมทั้งค่าใช้จ่ายที่เพิ่มขึ้นในการทำงานนั้นต่อให้แล้วเสร็จตามสัญญาซึ่งผู้ว่าจ้างจะ
</br>หักเอาจากจำนวนเงินใดๆ ที่จะจ่ายให้แก่ผู้รับจ้างก็ได้
</p>
<p class='tab3 t-16'>
    การที่ผู้ว่าจ้างไม่ใช้สิทธิเลิกสัญญาดังกล่าวตามวรรคหนึ่งไม่เป็นเหตุให้ผู้รับจ้างพ้นจากความ
</br>รับผิดตามสัญญา
</p>

<!-- ข้อ 11 การบังคับค่าปรับ ค่าเสียหาย และค่าใช้จ่าย -->
<p class='tab2 t-16'><b>ข้อ ๑๑ การบังคับค่าปรับ ค่าเสียหาย และค่าใช้จ่าย</b></p>
<p class='tab3 t-16'>
    ในกรณีที่ผู้รับจ้างไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่งด้วยเหตุใดๆ ก็ตาม จนเป็นเหตุให้เกิดค่า
</br>ปรับ ค่าเสียหาย หรือค่าใช้จ่ายแก่ผู้ว่าจ้าง ผู้รับจ้างต้องชดใช้ค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายดังกล่าวให้แก่ผู้ว่าจ้างโดยสิ้นเชิงภายในกำหนด{result.EnforcementOfFineDays} วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้าง หากผู้รับจ้างไม่
</br>ชดใช้ให้ถูกต้องครบถ้วนภายในระยะเวลาดังกล่าวให้ผู้ว่าจ้างมีสิทธิที่จะหักเอาจากจำนวนเงินค่าจ้าง 
</br>ที่ต้องชำระหรือบังคับจากหลักประกันการปฏิบัติตามสัญญาได้ทันที
</p>
<p class='tab3 t-16'>
    หากค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายที่บังคับจากจำนวนเงินค่าจ้างที่ต้องชำระ  หรือจากหลักประกันการปฏิบัติตามสัญญาแล้วยังไม่เพียงพอ ผู้รับจ้างยินยอมชำระส่วนที่เหลือที่ยังขาดอยู่จนครบถ้วน 
</br>ตามจำนวนค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายนั้น ภายในกำหนด{result.EnforcementOfFineDays} วัน นับถัดจากวันที่ได้รับแจ้งเป็น
</br>หนังสือจากผู้ว่าจ้าง
</p>

<!-- ข้อ 12 การงดหรือลดค่าปรับ หรือการขยายเวลาในการปฏิบัติตามสัญญา -->
<p class='tab2 t-16'><b>ข้อ ๑๒ การงดหรือลดค่าปรับ หรือการขยายเวลาในการปฏิบัติตามสัญญา</b></p>
<p class='tab3 t-16'>
    ในกรณีที่มีเหตุเกิดจากความผิดหรือความบกพร่องของฝ่ายผู้ว่าจ้าง หรือเหตุสุดวิสัย หรือเกิด
จากพฤติการณ์อันหนึ่งอันใดที่ผู้รับจ้างไม่ต้องรับผิดตามกฎหมาย หรือเหตุอื่นตามที่กำหนดในกฎกระทรวง 
</br>ซึ่งออกตามความในกฎหมายว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ ทำให้ผู้รับจ้างไม่สามารถ
</br>ปฏิบัติตามเงื่อนไขและกำหนดเวลาในข้อ ๕ ข้อ ๖ หรือข้อ ๗ ได้ ผู้รับจ้างจะต้องแจ้งเหตุหรือพฤติการณ์
</br>ดังกล่าวพร้อมหลักฐานเป็นหนังสือให้ผู้ว่าจ้างทราบ เพื่อของดหรือลดค่าปรับ หรือขยายเวลาทำการตามสัญญา
</br>ภายใน ๑๕ (สิบห้า) วันนับถัดจากวันที่เหตุนั้นสิ้นสุดลง หรือตามที่กำหนดในกฎกระทรวงดังกล่าว แล้วแต่กรณี
</p>
<p class='tab3 t-16'>
    ถ้าผู้รับจ้างไม่ปฏิบัติให้เป็นไปตามความในวรรคหนึ่ง ให้ถือว่าผู้รับจ้างได้สละสิทธิเรียกร้อง
</br>ในการที่จะของดหรือลดค่าปรับ หรือขยายเวลาทำการตามสัญญาโดยไม่มีเงื่อนไขใดๆทั้งสิ้น เว้นแต่กรณีเหตุ
</br>เกิดจากความผิดหรือความบกพร่องของฝ่ายผู้ว่าจ้างซึ่งมีหลักฐานชัดแจ้งหรือผู้ว่าจ้างทราบดีอยู่แล้วตั้งแต่ต้น
</p>
<p class='tab3 t-16'>
    การงดหรือลดค่าปรับ หรือขยายกำหนดเวลาทำการตามสัญญาตามวรรคหนึ่ง อยู่ในดุล
</br>พินิจของผู้ว่าจ้างที่จะพิจารณาตามที่เห็นสมควร
</p>

<!-- ลายเซ็นและพยาน -->
</br>
</br>
{signatoryHtml}

<!-- หน้าถัดไป: วิธีปฏิบัติเกี่ยวกับสัญญา -->
<div style='page-break-before: always;'></p>
<p class='text-center t-22' style='font-weight:bold;'>วิธีปฏิบัติเกี่ยวกับสัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์</p>
<p class='tab2 t-16'>(1) ให้ระบุเลขที่สัญญาในปีงบประมาณหนึ่งๆ ตามลำดับ</p>
<p class='tab2 t-16'>(2) ให้ระบุชื่อของหน่วยงานของรัฐที่เป็นนิติบุคคล เช่น กรม ก. หรือรัฐวิสาหกิจ ข. เป็นต้น</p>
<p class='tab2 t-16'>(3) ให้ระบุชื่อและตำแหน่งของหัวหน้าหน่วยงานของรัฐที่เป็นนิติบุคคลนั้น หรือผู้ที่ได้รับมอบอำนาจ เช่น นาย ก. อธิบดีกรม………...… หรือ นาย ข. ผู้ได้รับมอบอำนาจจากอธิบดีกรม………......………..</p>
<p class='tab2 t-16'>(4) ให้ระบุชื่อผู้รับจ้าง</p>
<p class='tab3 t-16'>ก. กรณีนิติบุคคล เช่น ห้างหุ้นส่วนสามัญจดทะเบียน ห้างหุ้นส่วนจำกัด บริษัทจำกัด</p>
<p class='tab3 t-16'>ข. กรณีบุคคลธรรมดา ให้ระบุชื่อและที่อยู่</p>
<p class='tab2 t-16'>(5) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
<p class='tab2 t-16'>(6) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
<p class='tab2 t-16'>(7) ให้กำหนดในอัตราระหว่างร้อยละ ๐.๐๒๕ – ๐.๐๓๕ ของราคาตามสัญญาต่อชั่วโมง</p>
<p class='tab2 t-16'>(8) อัตราค่าปรับตามสัญญาข้อ ๗ ให้กำหนดเป็นรายวันในอัตราระหว่างร้อยละ ๐.๐๑-๐.๑๐ ตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ ข้อ ๑๖๒ ส่วนกรณีจะปรับร้อยละเท่าใด ให้อยู่ในดุลพินิจของหน่วยงานของรัฐผู้ว่าจ้างที่จะพิจารณา โดยคำนึงถึงราคาและลักษณะของพัสดุที่จ้าง ซึ่งอาจมีผลกระทบต่อการที่ผู้รับจ้างจะหลีกเลี่ยงไม่ปฏิบัติตามสัญญา แต่ทั้งนี้การที่จะกำหนดค่าปรับเป็นร้อยละเท่าใดจะต้องกำหนดไว้ในเอกสารเชิญชวนด้วย</p>
<p class='tab2 t-16'>(9) “หลักประกัน” หมายถึง หลักประกันที่ผู้รับจ้างนำมามอบไว้แก่หน่วยงานของรัฐ เมื่อลงนามในสัญญา เพื่อเป็นการประกันความเสียหายที่อาจจะเกิดขึ้นจากการปฏิบัติตามสัญญา ดังนี้</p>
<p class='tab3 t-16'>(๑)เงินสด</p>
<p class='tab3 t-16'>(๒)เช็คหรือดราฟท์ที่ธนาคารเซ็นสั่งจ่าย ซึ่งเป็นเช็คหรือดราฟท์ลงวันที่ที่ใช้เช็คหรือดราฟท์นั้นชำระต่อเจ้าหน้าที่ หรือก่อนวันนั้นไม่เกิน ๓ วันทำการ</p>
<p class='tab3 t-16'>(๓)หนังสือคํ้าประกันของธนาคารภายในประเทศตามตัวอย่างที่คณะกรรมการนโยบายกําหนด โดยอาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้</p>
<p class='tab3 t-16'>(๔)หนังสือค้ำประกันของบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทยตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบ โดยอนุโลมให้ใช้ตามตัวอย่างหนังสือค้ำประกันของธนาคารที่คณะกรรมการนโยบายกําหนด</p>
<p class='tab3 t-16'>(๕)พันธบัตรรัฐบาลไทย</p>
<p class='tab2 t-16'>(10) ให้กำหนดจำนวนเงินหลักประกันการปฏิบัติตามสัญญาตามระเบียบการะทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ ข้อ ๑๖๘</p>
<p class='tab2 t-16'>(11) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
<p class='tab2 t-16'>(12) อัตราค่าปรับตามสัญญาข้อ ๙ กรณีผู้รับจ้างไปจ้างช่วงบางส่วนโดยไม่ได้รับอนุญาตจากผู้ว่าจ้างต้องกำหนดค่าปรับเป็นจำนวนเงินไม่น้อยกว่าร้อยละสิบของวงเงินของงานที่จ้างช่วงตามสัญญา</p>
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
        .contract, . {{
            margin: 12px 0;
            line-height: 1.7;
        }}
        . {{
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
        }
        catch (Exception ex)
        {
            throw new Exception("SLA 308-60 data not found.");
        }


    }
    #endregion 4.1.1.2.12.สัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ร.310-60

}
