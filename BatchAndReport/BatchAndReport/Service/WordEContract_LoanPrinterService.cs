using BatchAndReport.DAO;
using BatchAndReport.Entities;
using BatchAndReport.Models;
using BatchAndReport.Services;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using Spire.Doc.Documents;
using System.Text;
using System.Threading.Tasks;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;


public class WordEContract_LoanPrinterService
{
    private readonly WordServiceSetting _w;
    private readonly Econtract_Report_PMLDAO _pml;
    private readonly IConverter _pdfConverter; // Add this for PDF conversion
    private readonly E_ContractReportDAO _eContractReportDAO;
    private readonly EContractDAO _eContractDAO;
    public WordEContract_LoanPrinterService(WordServiceSetting ws, Econtract_Report_PMLDAO pml, IConverter pdfConverter, E_ContractReportDAO eContractReportDAO, EContractDAO eContractDAO)
    {
        _w = ws;
        _pml = pml;
        _pdfConverter = pdfConverter; // Initialize the PDF converter
        _eContractReportDAO = eContractReportDAO;
        _eContractDAO = eContractDAO;
    }
    #region 4.1.1.2.13.สัญญาเช่าเครื่องถ่ายเอกสาร ร.314-60
    public async Task<byte[]> OnGetWordContact_LoanPrinter(string id)
    {
        try {
        var result =await _pml.GetPMLAsync(id);
            if (result == null)
            {
                throw new Exception("Data not found for the given ID.");
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
                    body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("สัญญาเช่าเครื่องถ่ายเอกสาร", "000000", "36"));
                    // 2. Document title and subtitle
                    body.AppendChild(WordServiceSetting.EmptyParagraph());
                    body.AppendChild(WordServiceSetting.RightParagraph("สัญญาเลขที่ "+result.Contract_Number+""));



                    string datestring = CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now);
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("สัญญาฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.)\r\n"
                        + "ตำบล/แขวง ทุ่งสองห้อง อำเภอ/เขต หลักสี่\r\n" +
                    "จังหวัด กรุงเทพ เมื่อ" + datestring + "\r\n" +
                    "ระหว่าง " + result.Contract_Organization + "\r\n" +
                    "โดย " + result.SignatoryName + "\r\n" +
                    "ซึ่งต่อไปในสัญญานี้เรียกว่า “ผู้เช่า” ฝ่ายหนึ่ง กับ…" + result.ContractorName + "", null, "32"));

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
                        "สัญญานี้เรียกว่า “ผู้ให้เช่า” อีกฝ่ายหนึ่ง กับ…" + result.ContractorName + "", null, "32"));

                    }

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("คู่สัญญาได้ตกลงกันมีข้อความดังต่อไปนี้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 1 ข้อตกลงเช่า", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้เช่าตกลงเช่าและผู้ให้เช่าตกลงให้เช่าเครื่องถ่ายเอกสาร ยี่ห้อ "+result.RentalCopierBrand+" รุ่น"+result.RentalCopierModel+"" +
                    "หมายเลขเครื่อง "+result.RentalCopierNumber+" จำนวน "+result.RentalCopierAmount+" เครื่อง ซึ่งต่อไปในสัญญานี้เรียกว่า " +
                    "“เครื่องถ่ายเอกสารที่เช่า” เพื่อใช้ในกิจการของผู้เช่าตามเอกสารแนบท้ายสัญญา", null, "32"));

                    string strStartDate = CommonDAO.ToThaiDateStringCovert(result.RentalStartDate??DateTime.Now);
                    string strEndDate = CommonDAO.ToThaiDateStringCovert(result.RentalEndDate ?? DateTime.Now);
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("การเช่าเครื่องถ่ายเอกสารที่เช่าตามวรรคหนึ่งมีกำหนดระยะเวลา "+result.RentalYears+" ปี" +
                    " "+result.RentalMonths+" เดือน ตั้งแต่"+ strStartDate + " ถึง "+ strEndDate + "", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ให้เช่ารับรองว่าเครื่องถ่ายเอกสารที่เช่าตามสัญญานี้เป็นเครื่องถ่ายเอกสารใหม่ " +
                    "ที่ไม่เคยใช้งานมาก่อน ผู้ให้เช่าได้ชำระภาษี อากร ค่าธรรมเนียมต่างๆ ครบถ้วนถูกต้องตามกฎหมายแล้ว ผู้ให้เช่ามีสิทธินำมาให้เช่าโดยปราศจากการรอนสิทธิ ทั้งรับรองว่าเครื่องถ่ายเอกสารที่เช่ามีคุณสมบัติ คุณภาพ" +
                    "และคุณลักษณะไม่ต่ำกว่าที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก และผู้ให้เช่าได้ตรวจสอบแล้วว่าเครื่องถ่ายเอกสารที่เช่าตลอดจนอุปกรณ์ทั้งปวงปราศจากความชำรุดบกพร่อง", null, "32"));

                    string strPerUnit = CommonDAO.NumberToThaiText(result.RatePerUnit??0);
                    string strRateTotal = CommonDAO.NumberToThaiText(result.RateTotal ?? 0);
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(7)ข้อ 2ค่าเช่าเครื่องถ่ายเอกสาร", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้เช่าตกลงชำระค่าเช่าแก่ผู้ให้เช่าเป็นรายเดือนตามเดือนปฏิทินในอัตราค่าเช่าเดือนละ" +
                    " "+ result.RatePerUnit + " บาท ("+ strPerUnit + ") ต่อเครื่องถ่ายเอกสารหนึ่งเครื่อง รวมเป็นค่าเช่าทั้งสิ้นเดือนละ" +
                    ""+ result.RateTotal + "บาท ("+ strRateTotal + ") โดยประเมินจากจำนวนสำเนาเอกสารที่ถ่ายทั้งสิ้นเดือนละ" +
                    " "+result.EstCopiesPerMonth+" แผ่น", null, "32"));
                    body.AppendChild(WordServiceSetting.EmptyParagraph());
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หากเดือนใดจำนวนสำเนาเอกสารที่ผู้เช่าได้ถ่ายจากเครื่องถ่ายเอกสารที่เช่ามีจำนวนทั้งสิ้น ไม่ถึง "+result.IfNotCopiesAmount+" แผ่น การชำระค่าเช่าในเดือนนั้นให้เปลี่ยนเป็นคิดคำนวณจากจำนวนสำเนาเอกสารที่ถ่ายในเดือนนั้นๆ ในอัตราสำเนาแผ่นละ "+result.CopiesRate??0+" ", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("จำนวนสำเนาเอกสารที่ถ่ายตามความในสองวรรคก่อน ให้หมายความถึงสำเนาเอกสารที่ถ่ายออกมาโดยเรียบร้อยสมบูรณ์เท่านั้น การวินิจฉัยว่าสำเนาเอกสารแผ่นใดเป็นสำเนาเอกสารที่เรียบร้อยสมบูรณ์หรือเป็นสำเนาเอกสารเสีย ให้เป็นดุลพินิจของผู้เช่าหรือเจ้าหน้าที่ของผู้เช่า และการวินิจฉัยดังกล่าวให้เป็นที่สุด ผู้ให้เช่าจะโต้แย้งใดๆ มิได้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ค่าเช่าตามวรรคหนึ่งและวรรคสองได้รวมภาษีมูลค่าเพิ่ม ค่าใช้จ่ายในการบำรุงรักษาและซ่อมแซม ค่าตรวจสภาพ ค่าอะไหล่ ค่าวัสดุสิ้นเปลือง (ยกเว้นค่ากระดาษถ่ายเอกสาร) ไว้ด้วยแล้ว", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในการชำระค่าเช่า ผู้ให้เช่าต้องส่งใบแจ้งหนี้เรียกเก็บค่าเช่าเมื่อสิ้นเดือนแต่ละเดือน โดยผู้เช่าจะชำระค่าเช่าหลังจากที่ได้ตรวจสอบแล้วว่าถูกต้อง", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่การเช่าเดือนแรกและเดือนสุดท้ายเป็นการเช่าไม่เต็มเดือนปฏิทิน ให้ใช้วิธีการคำนวณค่าเช่าตามวรรคหนึ่งหรือวรรคสอง แล้วแต่กรณี แต่อัตราค่าเช่าตามวรรคหนึ่งให้คิดเป็นรายวัน ตามจำนวนวันที่เช่าจริง โดยคำนวณจากเดือนหนึ่งมี 30 (สามสิบ) วัน และให้ลดจำนวนสำเนาเอกสารที่ระบุตามวรรคสองลงตามสัดส่วนนั้น", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(8)การจ่ายเงินตามเงื่อนไขแห่งสัญญานี้ ผู้เช่าจะโอนเงินเข้าบัญชีเงินฝากธนาคารของผู้ให้เช่า ชื่อธนาคาร "+result.SaleBankName+ " สาขา "+result.SaleBankBranch+" ชื่อบัญชี " + result.SaleBankAccountName+" เลขที่บัญชี "+result.SaleBankAccountNumber+" ทั้งนี้ ผู้ให้เช่าตกลงเป็นผู้รับภาระเงินค่าธรรมเนียม หรือค่าบริการอื่นใดเกี่ยวกับการโอน รวมทั้งค่าใช้จ่ายอื่นใด (ถ้ามี) ที่ธนาคารเรียกเก็บ และยินยอมให้มีการหักเงินดังกล่าวจากจำนวนเงินโอนในงวดนั้นๆ (ความในวรรคนี้ใช้สำหรับกรณีที่หน่วยงานของรัฐจะจ่ายเงินตรงให้แก่ผู้ให้เช่า (ระบบ Direct Payment) โดยการโอนเงินเข้าบัญชีเงินฝากธนาคารของผู้ให้เช่าตามแนวทางที่กระทรวงการคลังหรือหน่วยงานของรัฐเจ้าของงบประมาณเป็นผู้กำหนด แล้วแต่กรณี)", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 3เอกสารอันเป็นส่วนหนึ่งของสัญญา", null, "32", true));

                    // select request
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เอกสารแนบท้ายสัญญาดังต่อไปนี้ให้ถือเป็นส่วนหนึ่งของสัญญานี้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("3.1 ผนวก 1 ………………….(ใบเสนอราคา)…………...............จำนวน.....(…….) หน้า", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("3.2 ผนวก 2 ….(แค็ตตาล็อก คุณลักษณะและรายละเอียดจำนวน.....(…….) หน้า  ของเครื่องถ่ายเอกสารที่เช่า).......", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("3.3 ผนวก 3 ……………(กำหนดการบำรุงรักษา)…………......จำนวน.....(…….) หน้า", null, "32"));
                    // select request

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ความใดในเอกสารแนบท้ายสัญญาที่ขัดหรือแย้งกับข้อความในสัญญานี้ ให้ใช้ข้อความในสัญญานี้บังคับ และในกรณีที่เอกสารแนบท้ายสัญญาขัดแย้งกันเอง ผู้ให้เช่าจะต้องปฏิบัติตามคำวินิจฉัยของผู้เช่า คำวินิจฉัยของผู้เช่าให้ถือเป็นที่สุด และผู้ให้เช่าไม่มีสิทธิเรียกร้องค่าเช่า ค่าเสียหาย หรือค่าใช้จ่ายใดๆ เพิ่มเติมจากผู้เช่าทั้งสิ้น", null, "32"));


                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 4การส่งมอบ", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ให้เช่าต้องส่งมอบและติดตั้งเครื่องถ่ายเอกสารที่เช่าตามสัญญานี้ ให้ถูกต้องครบถ้วนตามสัญญานี้ " +
                        "ในลักษณะพร้อมใช้งานได้ตามที่กำหนด ณ ........................ ภายในวันที่.......................... ซึ่งผู้ให้เช่าเป็นผู้จัดหาอุปกรณ์ประกอบ " +
                        "พร้อมทั้งเครื่องมือที่จำเป็นในการติดตั้งและใช้งาน โดยผู้ให้เช่าเป็นผู้ออกค่าใช้จ่ายเองทั้งสิ้น", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ทั้งนี้ ผู้ให้เช่าต้องแจ้งเวลาติดตั้งแล้วเสร็จพร้อมที่จะใช้งานและส่งมอบเครื่องได้เป็นหนังสือ" +
                        "ต่อผู้เช่า ณ "+result.DeliveryLocation+" ในวันและเวลาทำการของผู้เช่าก่อนวันกำหนดส่งมอบตามวรรคหนึ่งไม่น้อยกว่า "+result.DeliveryType+" วันทำการของผู้เช่า", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในการส่งมอบตามวรรคหนึ่ง ผู้ให้เช่าต้องส่งพนักงานมาดำเนินการทดสอบประสิทธิภาพและ" +
                        "แนะนำวิธีการใช้เครื่องให้คณะกรรมการตรวจรับพัสดุได้พิจารณาตามรายละเอียดคุณลักษณะเฉพาะที่ระบุไว้ในข้อ 1 และสำเนาที่ถ่ายจะต้องมีความชัดเจนสะอาด" +
                        "ไม่มีรอยหมึกเปื้อนตามส่วนต่างๆ โดยในการนี้ผู้ให้เช่าไม่คิดค่าใช้จ่ายใดๆ จากผู้เช่าทั้งสิ้น", null, "32"));


                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 5การตรวจรับ", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เมื่อผู้เช่าได้ตรวจรับเครื่องถ่ายเอกสารที่ส่งมอบตามข้อ4 และเห็นว่าถูกต้องครบถ้วนตามสัญญานี้แล้ว ผู้เช่าจะออกหลักฐานการรับมอบเครื่องถ่ายเอกสารที่เช่าไว้เป็นหนังสือ เพื่อผู้ให้เช่านำมาใช้เป็นหลักฐานประกอบการขอรับเงินค่าเช่า", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในการตรวจรับเครื่องถ่ายเอกสารที่ส่งมอบตามวรรคหนึ่ง ถ้าปรากฏว่าเครื่องถ่ายเอกสาร" +
                        "ซึ่งผู้ให้เช่าส่งมอบไม่ถูกต้องครบถ้วนตามสัญญา หรือติดตั้งและส่งมอบถูกต้องครบถ้วนภายในกำหนด" +
                        "แต่ไม่สามารถใช้งานได้อย่างครบถ้วนและมีประสิทธิภาพตามสัญญา ผู้เช่าทรงไว้ซึ่งสิทธิที่จะไม่รับเครื่องถ่าย" +
                        "เอกสารนั้น ในกรณีเช่นว่านี้ ผู้ให้เช่าต้องรีบนำเครื่องถ่ายเอกสารนั้นกลับคืนไปทันที และต้องนำเครื่องถ่าย" +
                        "เอกสารเครื่องใหม่ที่มีคุณสมบัติเดียวกัน หรือไม่ต่ำกว่าเครื่องถ่ายเอกสารที่กำหนดไว้ในสัญญานี้ มาส่งมอบให้ใหม่ " +
                        "ภายใน "+result.TotalDay+" วัน ด้วยค่าใช้จ่ายของผู้ให้เช่าเองทั้งสิ้น และระยะเวลาที่เสียไปเพราะเหตุดังกล่าว" +
                        " ผู้ให้เช่าจะนำมาอ้างเป็นเหตุของดหรือลดค่าปรับหรือขยายเวลาส่งมอบไม่ได้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หากผู้ให้เช่าไม่นำเครื่องถ่ายเอกสารที่ส่งมอบไม่ถูกต้องกลับคืนไปในทันทีดังกล่าวในวรรคสอง และเกิดความเสียหายแก่เครื่องถ่ายเอกสารนั้น ผู้เช่าไม่ต้องรับผิดชอบในความเสียหายดังกล่าว", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้ให้เช่าส่งมอบเครื่องถ่ายเอกสารที่เช่าถูกต้องแต่ไม่ครบจำนวน หรือส่งมอบครบจำนวนแต่ไม่ถูกต้องทั้งหมด ผู้เช่ามีสิทธิจะรับมอบเฉพาะส่วนที่ถูกต้อง โดยออกหลักฐานการรับมอบเฉพาะส่วนนั้นก็ได้ ในกรณีเช่นนี้ผู้เช่าจะชำระค่าเช่าเฉพาะเครื่องถ่ายเอกสารที่เช่าที่รับมอบไว้", null, "32"));


                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 6การงดหรือลดค่าปรับ หรือขยายเวลาในการปฏิบัติตามสัญญา", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้ให้เช่าไม่สามารถส่งมอบเครื่องถ่ายเอกสารที่เช่าให้แก่ผู้เช่าได้โดยครบถ้วนถูกต้องภายในกำหนดเวลาตามสัญญา หรือถ้าผู้ให้เช่าไม่ดำเนินการหรือ" +
                        "ไม่สามารถซ่อมแซมแก้ไขเครื่องถ่ายเอกสารที่เช่าภายในระยะเวลาตามข้อ 8.2 และผู้ให้เช่าไม่จัดหาเครื่องถ่ายเอกสารให้ผู้เช่าใช้แทนตามข้อ 8.3 อันเนื่อง" +
                        "มาจากเหตุสุดวิสัย หรือเหตุใดๆ อันเนื่องมาจากความผิดหรือความบกพร่องของฝ่ายผู้เช่าหรือจากพฤติการณ์อันหนึ่งอันใดที่ผู้ให้เช่าไม่ต้องรับ" +
                        "ผิดตามกฎหมาย หรือเหตุอื่นตามที่กำหนดในกฎกระทรวง ซึ่งออกตามความในกฎหมายว่าด้วยการจัดซื้อจัดจ้างและ" +
                        "การบริหารพัสดุภาครัฐ ผู้ให้เช่ามีสิทธิของดหรือลดค่าปรับหรือขยายกำหนดเวลาทำการตามสัญญาดังกล่าว โดยจะต้องแจ้งเหตุหรือพฤติการณ์ดัง" +
                        "กล่าวพร้อมหลักฐานเป็นหนังสือให้ผู้เช่าทราบภายใน ๑๕ (สิบห้า) วัน นับถัดจากวันที่เหตุนั้นสิ้นสุดลงหรือตามที่กำหนดในกฎกระทรวงดังกล่าว แล้วแต่กรณี", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ถ้าผู้ให้เช่าไม่ปฏิบัติให้เป็นไปตามความในวรรคหนึ่ง ให้ถือว่าผู้ให้เช่าได้สละสิทธิเรียกร้อง ในการที่จะของด" +
                        "หรือลดค่าปรับหรือขยายเวลาทำการตามสัญญาโดยไม่มีเงื่อนไขใดๆ ทั้งสิ้น เว้นแต่กรณีเหตุเกิดจากความผิดหรือความบกพร่องของฝ่ายผู้เช่าซึ่งมีหลักฐานชัดแจ้ง หรือผู้เช่าทราบดีอยู่แล้วตั้งแต่ต้น", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("การงดหรือลดค่าปรับหรือขยายกำหนดเวลาทำการตามสัญญาตามวรรคหนึ่ง อยู่ในดุลพินิจของผู้เช่าที่จะพิจารณาตามที่เห็นสมควร", null, "32"));


                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 7การบำรุงรักษาตรวจสภาพและซ่อมแซมเครื่องถ่ายเอกสารที่เช่า", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ผู้ให้เช่ามีหน้าที่บำรุงรักษาเครื่องถ่ายเอกสารที่เช่าให้อยู่ในสภาพใช้งานได้ดีอยู่เสมอด้วยค่าใช้จ่ายของผู้ให้เช่า โดยต้อง" +
                        "จัดหาช่างผู้มีความรู้ ความชำนาญ และฝีมือดีมาตรวจสอบ บำรุงรักษาและซ่อมแซมแก้ไขเครื่องถ่ายเอกสารที่เช่าตลอดอายุสัญญาเช่านี้ อย่างน้อยเดือนละ "+result.MaintenancePermonth+" ครั้ง โดยให้มีระยะเวลาห่างกันไม่น้อยกว่า "+result.MaintenanceInterval+" วัน", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("สิ่งของที่ใช้สิ้นเปลืองทุกชนิดรวมทั้งอะไหล่ ยกเว้นกระดาษสำหรับถ่ายเอกสาร ผู้ให้เช่าจะเป็นผู้จัดส่งให้โดยไม่คิดมูลค่า โดยที่ผู้ให้เช่าจะจัดให้มีไว้ในความครอบครองของผู้เช่าให้เพียงพออยู่เสมอ อุปกรณ์สิ้นเปลืองดังกล่าว เช่น ลูกโม่ถ่ายภาพ ผงหมึก ผงประจุภาพ หมึกพิมพ์ วัสดุที่ใช้ทำความสะอาดถุงกรอง แปรง น้ำมันหล่อลื่น และอุปกรณ์อื่นๆ ที่จำเป็นเพื่อให้เครื่องถ่ายเอกสารใช้งานได้ตามปกติตลอดเวลา", null, "32"));


                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 8หน้าที่ของผู้ให้เช่า", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("8.1 ผู้ให้เช่ามีหน้าที่ฝึกอบรมวิธีใช้เครื่องถ่ายเอกสารที่เช่าให้แก่เจ้าหน้าที่ของผู้เช่า จนสามารถใช้งานเครื่องถ่ายเอกสารได้ และผู้ให้เช่าตกลงจะฝึกอบรมวิธีการใช้เครื่องถ่ายเอกสารที่เช่าให้แก่เจ้าหน้าที่ของผู้เช่าทุกครั้ง หากผู้เช่าร้องขอโดยเหตุที่มีการเปลี่ยนแปลงโยกย้ายเจ้าหน้าที่ของผู้เช่าและเจ้าหน้าที่คนนั้นยังไม่เคยได้รับการฝึกอบรมมาก่อนโดยผู้ให้เช่าเป็นผู้รับผิดชอบค่าใช้จ่ายในการฝึกอบรมทั้งสิ้น", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("8.2 ในกรณีเครื่องถ่ายเอกสารที่เช่าชำรุดบกพร่องหรือขัดข้องใช้งานไม่ได้ตามปกติ ผู้ให้เช่าจะต้องจัดให้ช่างที่มีความรู้ความชำนาญและฝีมือดีมาจัดการซ่อมแซมแก้ไขให้อยู่ในสภาพใช้งานได้ดีตามปกติ โดยผู้ให้เช่าจะต้องเริ่มจัดการซ่อมแซมแก้ไขในทันทีที่ได้รับแจ้งจากผู้เช่าหรือผู้ที่ได้รับมอบหมายจากผู้เช่าแล้ว และให้แล้วเสร็จใช้งานได้ดีดังเดิมอย่างช้าต้องไม่เกิน "+result.CopierFixDays+"  ชั่วโมง ตั้งแต่เวลาที่ได้รับแจ้ง", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("8.3 ในกรณีที่เครื่องถ่ายเอกสารที่เช่ามีความชำรุดบกพร่องหรือขัดข้องใช้งานไม่ได้ตามปกติ และการซ่อมแซมต้องใช้เวลาเกินกว่า "+result.ReplaceFixDays+" ชั่วโมง ตามที่กำหนดในข้อ 8.2 หรือไม่อาจซ่อมแซมแก้ไขให้ดีได้ดังเดิม ผู้ให้เช่าต้องจัดหาเครื่องถ่ายเอกสารที่มีคุณสมบัติ คุณภาพ ความสามารถ และประสิทธิภาพในการใช้งานไม่ต่ำกว่าของเครื่องเดิมมาให้ผู้เช่าใช้แทนทันที ", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 9ค่าปรับกรณีความชำรุดบกพร่องของเครื่องถ่ายเอกสาร", null, "32"));

                    string strFinePerDays = CommonDAO.NumberToThaiText(result.FinePerDays??0);
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ถ้าผู้ให้เช่าไม่ดำเนินการหรือไม่สามารถซ่อมแซมแก้ไขเครื่องถ่ายเอกสารที่เช่าภายในระยะเวลาตามข้อ " +
                        "8.2 และผู้ให้เช่าไม่จัดหาเครื่องถ่ายเอกสารให้ผู้เช่าใช้แทนตามข้อ 8.3 ผู้ให้เช่ายินยอมให้ผู้เช่าปรับเป็นรายวัน ในอัตราวันละ "+result.FinePerDays+" บาท ("+ strFinePerDays + ") ต่อเครื่อง ตั้งแต่พ้นกำหนดระยะเวลาตามข้อ " +
                        "8.2 จนถึงวันที่ผู้ให้เช่าซ่อมแซมแก้ไขให้อยู่ในสภาพใช้งานได้ดีตามปกติ หรือผู้ให้เช่าจัดหาเครื่องถ่ายเอกสารมาให้ผู้เช่าใช้งานแทน หรือจนกว่าผู้เช่าจะใช้สิทธิบอกเลิกสัญญา " +
                        "ทั้งนี้ ผู้เช่าไม่ต้องจ่าย  ค่าเช่าในระหว่างเวลาที่ผู้เช่าไม่สามารถใช้เครื่องถ่ายเอกสารที่เช่าตามสัญญานี้ โดยยินยอมให้ผู้เช่าหักค่าปรับดังกล่าวออกจากค่าเช่าตามข้อ 2 หรือบังคับเอาจากหลักประกันตามข้อ 10 ก็ได้", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 10หลักประกันการปฏิบัติตามสัญญา", null, "32", true));

                    string strGuaranteeAmount = CommonDAO.NumberToThaiText(result.GuaranteeAmount ?? 0);
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในขณะทำสัญญานี้ผู้ให้เช่าได้นำหลักประกันเป็น "+result.GuaranteeType+" เป็นจำนวนเงิน "+result.GuaranteeAmount+" บาท ("+ strGuaranteeAmount + ") ซึ่งเท่ากับร้อยละ "+result.GuaranteePercent+" ของค่าเช่าทั้งหมดตามสัญญา มามอบให้แก่ผู้เช่าเพื่อเป็นหลักประกันการปฏิบัติตามสัญญานี้ ", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(14) กรณีผู้ให้เช่าใช้หนังสือค้ำประกันมาเป็นหลักประกันการปฏิบัติตามสัญญา หนังสือค้ำประกันดังกล่าวจะต้องออกโดยธนาคารที่ประกอบกิจการในประเทศไทย หรือโดยบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบตามแบบที่คณะกรรมการนโยบายการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐกำหนด หรืออาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้ และจะต้องมีอายุการค้ำประกันตลอดไปจนกว่าผู้ให้เช่าพ้นข้อผูกพันตามสัญญานี้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หลักประกันที่ผู้ให้เช่านำมามอบให้ตามวรรคหนึ่ง จะต้องมีอายุครอบคลุมความรับผิด ทั้งปวงของผู้ให้เช่าตลอดอายุสัญญา ถ้าหลักประกันที่ผู้ให้เช่านำมามอบให้ดังกล่าวลดลงหรือเสื่อมค่าลง หรือมีอายุไม่ครอบคลุมถึงความรับผิดของผู้ให้เช่าตลอดอายุสัญญา ไม่ว่าด้วยเหตุใดๆ ก็ตาม รวมถึงกรณีผู้ให้เช่าส่งมอบและติดตั้งเครื่องถ่ายเอกสารล่าช้าเป็นเหตุให้ระยะเวลาการเช่าตามสัญญาเปลี่ยนแปลงไป ผู้ให้เช่าต้องหาหลักประกันใหม่หรือหลักประกันเพิ่มเติมให้มีจำนวนครบถ้วนตามวรรคหนึ่งมามอบให้แก่ผู้เช่าภายใน "+result.NewGuaranteeDays+" วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้เช่า", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หลักประกันที่ผู้ให้เช่านำมามอบไว้ตามข้อนี้ ผู้เช่าจะคืนให้แก่ผู้ให้เช่าโดยไม่มีดอกเบี้ยเมื่อผู้ให้เช่าพ้นจากข้อผูกพันและความรับผิดทั้งปวงตามสัญญานี้แล้ว", null, "32"));


                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 11การบอกเลิกสัญญา", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เมื่อครบกำหนดส่งมอบเครื่องถ่ายเอกสารที่เช่าตามสัญญาแล้ว ถ้าผู้ให้เช่าไม่ส่งมอบเครื่องถ่ายเอกสารที่เช่า หรือส่งมอบแต่เพียงบางส่วนให้แก่ผู้เช่า หรือส่งมอบเครื่องถ่ายเอกสารที่เช่าไม่ตรงตามสัญญาหรือมีคุณลักษณะเฉพาะไม่ถูกต้องตามสัญญา " +
                        "หรือส่งมอบแล้วเสร็จภายในกำหนดแต่ไม่สามารถใช้งานได้อย่างมีประสิทธิภาพหรือใช้งานได้ไม่ครบถ้วนตามสัญญา " +
                        "หรือผู้ให้เช่าไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่ง ผู้เช่ามีสิทธิบอกเลิกสัญญาทั้งหมดหรือแต่บางส่วนได้ การใช้สิทธิบอกเลิกสัญญานั้นไม่กระทบสิทธิของผู้เช่าที่จะเรียกร้องค่าเสียหายจากผู้ให้เช่า", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้เช่าใช้สิทธิบอกเลิกสัญญา ผู้เช่ามีสิทธิริบหรือบังคับจากหลักประกัน ตามข้อ 10 เป็นจำนวนเงินทั้งหมดหรือแต่บางส่วนก็ได้แล้วแต่ผู้เช่าจะเห็นสมควร และถ้าผู้เช่าต้องเช่า " +
                        "เครื่องถ่ายเอกสารจากบุคคลอื่นทั้งหมดหรือแต่บางส่วนภายในกำหนด "+result.TeminationReplaceDays+" เดือน นับถัดจากวันบอกเลิกสัญญา " +
                        "ผู้ให้เช่ายอมรับผิดชดใช้ค่าเช่าที่เพิ่มขึ้นจากค่าเช่าที่กำหนดไว้ในสัญญานี้ รวมทั้งค่าใช้จ่ายใดๆ ที่ผู้เช่าต้องใช้จ่ายในการจัดหาผู้ให้เช่าเครื่องถ่ายเอกสารที่เช่ารายใหม่ดังกล่าวด้วย", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีมีความจำเป็น ผู้เช่ามีสิทธิที่จะบอกเลิกสัญญาเช่านี้ก่อนครบกำหนดระยะเวลาการเช่าได้ โดยแจ้งเป็นหนังสือให้ผู้ให้เช่าทราบล่วงหน้าไม่น้อยกว่า ๓๐ (สามสิบ) วัน โดยผู้ให้เช่าจะไม่มีสิทธิเรียกร้องค่าเสียหายใดๆ จากผู้เช่า", null, "32"));



                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 12ค่าปรับกรณีส่งมอบล่าช้า", null, "32"));

                    string strLate = CommonDAO.NumberToThaiText(result.LateFinePerDays??0);
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้ให้เช่าส่งมอบเครื่องถ่ายเอกสารที่เช่าล่วงเลยกำหนดส่งมอบตามข้อ 4 และผู้เช่ามิได้ใช้สิทธิบอกเลิกสัญญาตามข้อ 11 วรรคหนึ่ง ผู้ให้เช่าจะต้องชำระค่าปรับให้ผู้เช่าเป็นรายวัน " +
                        "สำหรับเครื่องถ่ายเอกสารที่ยังไม่ได้ส่งมอบตามสัญญา ในอัตราวันละ "+result.LateFinePerDays+" บาท ("+ strLate + ") ต่อเครื่อง นับถัดจากวันที่ครบกำหนดส่งมอบตามสัญญาจนถึงวันที่ผู้ให้เช่าได้นำเครื่องถ่ายเอกสารที่เช่ามาส่งมอบให้แก่ผู้เช่าจนถูกต้องครบถ้วน", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในระหว่างที่ผู้เช่ายังมิได้ใช้สิทธิบอกเลิกสัญญานั้น หากผู้เช่าเห็นว่าผู้ให้เช่าไม่อาจปฏิบัติตามสัญญาต่อไปได้ " +
                        "ผู้เช่าจะใช้สิทธิบอกเลิกสัญญา และบังคับจากหลักประกันการปฏิบัติตามสัญญาตามข้อ ๑๐ กับเรียกร้องให้ชดใช้ค่าเช่าที่เพิ่มขึ้นตามที่กำหนดไว้ในข้อ ๑๑ วรรคสองก็ได้ " +
                        "และถ้าผู้เช่าได้แจ้งข้อเรียกร้องให้ชำระค่าปรับไปยังผู้ให้เช่าเมื่อครบกำหนดส่งมอบดังกล่าวแล้ว ผู้เช่ามีสิทธิที่จะปรับผู้ให้เช่าจนถึงวันบอกเลิกสัญญาได้อีกด้วย", null, "32"));


                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 13การบังคับค่าปรับ ค่าเสียหาย และค่าใช้จ่าย", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในกรณีที่ผู้ให้เช่าไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่งด้วยเหตุใดๆ ก็ตาม จนเป็นเหตุให้เกิดค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายแก่ผู้เช่า ผู้ให้เช่าต้องชดใช้ค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายดังกล่าวให้แก่ผู้เช่าโดยสิ้นเชิงภายในกำหนด "+result.EnforcementOfFineDays+" วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้เช่า หากผู้ให้เช่าไม่ชดใช้ให้ถูกต้องครบถ้วนภายในระยะเวลาดังกล่าวให้ผู้เช่ามีสิทธิที่จะหักเอาจากค่าเช่าที่ต้องชำระหรือบังคับจากหลักประกันการปฏิบัติตามสัญญาได้ทันที", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("หากค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายที่บังคับจากค่าเช่าที่ต้องชำระ หรือหลักประกันการปฏิบัติตามสัญญาแล้วยังไม่เพียงพอ ผู้ให้เช่ายินยอมชำระส่วนที่เหลือที่ยังขาดอยู่จนครบถ้วนตามจำนวนค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายนั้น ภายในกำหนด "+result.OutstandingPeriodDays+" วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้เช่า", null, "32"));


                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 14การโอนสิทธิของผู้ให้เช่า", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ในระหว่างอายุสัญญาเช่า ห้ามผู้ให้เช่าโอนสิทธิหน้าที่ตามสัญญาหรือกรรมสิทธิ์ในเครื่องถ่ายเอกสารที่เช่าแก่บุคคลอื่น เว้นแต่จะได้รับความยินยอมเป็นหนังสือจากผู้เช่าก่อน", null, "32"));

                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 15การนำเครื่องถ่ายเอกสารที่เช่ากลับคืนเมื่อสัญญาสิ้นสุดลง", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เมื่อสัญญาสิ้นสุดลงไม่ว่าจะเป็นการบอกเลิกสัญญาหรือครบกำหนดเวลาตามสัญญา ผู้ให้เช่าต้องนำเครื่องถ่ายเอกสารที่เช่ากลับคืนไปภายใน "+result.CopierSendBackDays+" วัน โดยผู้ให้เช่าเป็นผู้เสียค่าใช้จ่ายเองทั้งสิ้น", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ถ้าผู้ให้เช่าไม่นำเครื่องถ่ายเอกสารที่เช่ากลับคืนไปภายในกำหนดเวลาตามวรรคหนึ่ง ผู้เช่าไม่ต้องรับผิดชอบในความเสียหายใดๆ ทั้งสิ้นที่เกิดแก่เครื่องถ่ายเอกสารที่เช่าอันมิใช่ความผิดของผู้เช่า", null, "32"));


                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อ 16ข้อจำกัดความรับผิดของผู้เช่า", null, "32", true));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("เช่าไม่ต้องรับผิดในความเสียหายหรือสูญหายเมื่อเกิดอัคคีภัยหรือภัยพิบัติใดๆ หรือการโจรกรรมเครื่องถ่ายเอกสารที่เช่าตลอดจนการสูญหายหรือความเสียหายใดๆ ที่เกิดขึ้นแก่เครื่องถ่ายเอกสาร ที่เช่าอันไม่ใช่เกิดจากความผิดของผู้เช่าตลอดระยะเวลาที่เครื่องถ่ายเอกสารอยู่ในความครอบครองของผู้เช่า", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("สัญญานี้ทำขึ้นสองฉบับมีข้อความถูกต้องตรงกัน คู่สัญญาได้อ่านและเข้าใจข้อความโดยละเอียดตลอดแล้ว จึงได้ลงลายมือชื่อ พร้อมประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยาน และคู่สัญญาต่างยึดถือไว้คนละหนึ่งฉบับ", null, "32"));

                    body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ "+result.OSMEP_Signer+"ผู้ว่าจ้าง"));
                    body.AppendChild(WordServiceSetting.CenteredParagraph("(................................................................................)"));
                    body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ "+result.Contract_Signer+"ผู้ว่าจ้าง"));
                    body.AppendChild(WordServiceSetting.CenteredParagraph("(................................................................................)"));
                    body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ "+result.OSMEP_Witness+"พยาน"));
                    body.AppendChild(WordServiceSetting.CenteredParagraph("(...............................................................................)"));
                    body.AppendChild(WordServiceSetting.CenteredParagraph("ลงชื่อ"+result.Contract_Witness+"พยาน"));
                    body.AppendChild(WordServiceSetting.CenteredParagraph("(...............................................................................)"));

                    // next page
                    body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                    body.AppendChild(WordServiceSetting.CenteredBoldColoredParagraph("วิธีปฏิบัติเกี่ยวกับสัญญาเช่าเครื่องถ่ายเอกสาร", "000000", "36"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(1) ให้ระบุเลขที่สัญญาในปีงบประมาณหนึ่งๆ ตามลำดับ", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(2) ให้ระบุชื่อของหน่วยงานของรัฐที่เป็นนิติบุคคล เช่น กรม ก. หรือรัฐวิสาหกิจ ข. เป็นต้น", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(3) ให้ระบุชื่อและตำแหน่งของหัวหน้าหน่วยงานของรัฐที่เป็นนิติบุคคลนั้น หรือผู้ที่ได้รับมอบอำนาจ เช่น นาย ก. อธิบดีกรม………...… หรือ นาย ข. ผู้ได้รับมอบอำนาจจากอธิบดีกรม………......………..", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(4) ให้ระบุชื่อผู้ให้เช่า", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ก. กรณีนิติบุคคล เช่น ห้างหุ้นส่วนสามัญจดทะเบียน ห้างหุ้นส่วนจำกัด บริษัทจำกัด", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("ข. กรณีบุคคลธรรมดา ให้ระบุชื่อและที่อยู่", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(5) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(6) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(7) หน่วยงานของรัฐอาจกำหนดเงื่อนไขการจ่ายค่าเช่าให้แตกต่างไปจากแบบสัญญาที่กำหนดได้ตามความเหมาะสมและจำเป็นและไม่ทำให้หน่วยงานของรัฐเสียเปรียบ หากหน่วยงานของรัฐเห็นว่าจะมีปัญหาในทางเสียเปรียบหรือไม่รัดกุมพอ ก็ให้ส่งร่างสัญญานั้นไปให้สำนักงานอัยการสูงสุดพิจารณาให้ความเห็นชอบก่อน", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(8) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(9) ชื่อสถานที่หน่วยงานของรัฐ", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(10) ให้พิจารณาถึงความจำเป็นและเหมาะสมของการใช้งาน", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(11) อัตราค่าปรับตามสัญญาข้อ 9 ให้กำหนดเป็นรายวันในอัตราระหว่างร้อยละ 0.01 – 0.20 ตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. 2560 ข้อ 162 ส่วนกรณีจะปรับร้อยละเท่าใด ให้อยู่ในดุลยพินิจของหน่วยงานของรัฐผู้เช่าที่จะพิจารณา โดยคำนึงถึงราคาและลักษณะของพัสดุที่เช่า ซึ่งอาจมีผลกระทบต่อการที่ผู้ให้เช่าจะหลีกเลี่ยงไม่ปฏิบัติตามสัญญา แต่ทั้งนี้การที่จะกำหนดค่าปรับเป็นร้อยละเท่าใดจะต้องกำหนดไว้ในเอกสารเชิญชวนด้วย", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(12) “หลักประกัน” หมายถึง หลักประกันที่ผู้ให้เช่านำมามอบไว้แก่หน่วยงานของรัฐ เมื่อลงนามในสัญญา เพื่อเป็นการประกันความเสียหายที่อาจจะเกิดขึ้นจากการปฏิบัติตามสัญญา ดังนี้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๑) เงินสด", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๒) เช็คหรือดราฟท์ ที่ธนาคารเซ็นสั่งจ่าย ซึ่งเป็นเช็คหรือดราฟท์ลงวันที่ที่ใช้เช็คหรือดราฟท์นั้นชำระต่อเจ้าหน้าที่ หรือก่อนวันนั้นไม่เกิน ๓ วันทำการ ", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๓) หนังสือคํ้าประกันของธนาคารภายในประเทศตามตัวอย่างที่คณะกรรมการนโยบายกําหนด โดยอาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๔) หนังสือค้ำประกันของบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบ โดยอนุโลมให้ใช้ตามตัวอย่างหนังสือค้ำประกันของธนาคารที่คณะกรรมการนโยบายกำหนด", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(๕) พันธบัตรรัฐบาลไทย", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(13) ให้กำหนดจำนวนเงินหลักประกันการปฏิบัติตามสัญญาตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. 2560 ข้อ 168", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(14) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("(15) กำหนดระยะเวลาตามความเหมาะสม เช่น 3 เดือน", null, "32"));
                    body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("(16) อัตราค่าปรับตามสัญญาข้อ 12 ให้กำหนดเป็นรายวันในอัตราระหว่างร้อยละ 0.01 – 0.20 ตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. 2560 ข้อ 162 ส่วนกรณีจะปรับร้อยละเท่าใด ให้อยู่ในดุลยพินิจของหน่วยงานของรัฐผู้เช่าที่จะพิจารณา โดยคำนึงถึงราคาและลักษณะของพัสดุที่เช่า ซึ่งอาจมีผลกระทบต่อการที่ผู้ให้เช่าจะหลีกเลี่ยงไม่ปฏิบัติตามสัญญา แต่ทั้งนี้การที่จะกำหนดค่าปรับเป็นร้อยละเท่าใดจะต้องกำหนดไว้ในเอกสารเชิญชวนด้วย ", null, "32"));


                    body.AppendChild(WordServiceSetting.EmptyParagraph());




                    WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);

                }
                stream.Position = 0;
                return stream.ToArray();
            }
        }
        catch (Exception ex)
        {
            // Log the exception if necessary
            throw new Exception("Error generating Word document", ex);
        }
      
    }
    #endregion 4.1.1.2.13.สัญญาเช่าเครื่องถ่ายเอกสาร ร.314-60


    public async Task<string> OnGetWordContact_LoanPrinter_ToPDF(string id,string typeContact)
    {
        try {

            var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");
            var cssPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "css", "contract.css").Replace("\\", "/");
            var result = await _pml.GetPMLAsync(id);
            if (result == null)
                throw new Exception("Data not found for the given ID.");

            string datestring = CommonDAO.ToThaiDateStringCovert(result.ContractSignDate ?? DateTime.Now);
            string strStartDate = CommonDAO.ToThaiDateStringCovert(result.RentalStartDate ?? DateTime.Now);
            string strEndDate = CommonDAO.ToThaiDateStringCovert(result.RentalEndDate ?? DateTime.Now);
            string strPerUnit = CommonDAO.NumberToThaiText(result.RatePerUnit ?? 0);
            string strRateTotal = CommonDAO.NumberToThaiText(result.RateTotal ?? 0);
            string strFinePerDays = CommonDAO.NumberToThaiText(result.FinePerDays ?? 0);
            string strGuaranteeAmount = CommonDAO.NumberToThaiText(result.GuaranteeAmount ?? 0);
            string strLate = CommonDAO.NumberToThaiText(result.LateFinePerDays ?? 0);

            string strContractorAuthDate = CommonDAO.ToThaiDateStringCovert(result.ContractorAuthDate??DateTime.Now);
            var listDocAtt = await _eContractDAO.GetRelatedDocumentsAsync(id, "PML31460");
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
    <div class='text-center t-22'> <b>แบบสัญญา</b> </div>
    <div class='text-center t-22'><b>สัญญาเช่าเครื่องถ่ายเอกสาร</b></div>

</br>

    <p class='text-right t-16'>สัญญาเลขที่ {result.Contract_Number}</p>
  <p class='tab3 t-16'>สัญญาฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.) </br>ตำบล/แขวง 
ทุ่งสองห้อง อำเภอ/เขต หลักสี่ จังหวัด กรุงเทพ เมื่อ {datestring} ระหว่าง {result.Contract_Organization} โดย {result.SignatoryName} ตำแหน่ง {result.SignatoryPosition} ซึ่งต่อไปในสัญญานี้เรียกว่า “ผู้เช่า” ฝ่ายหนึ่ง กับ {result.ContractorName}</p>
    {(result.ContractorType == "นิติบุคคล"
                ? $"<p class='t-16'>ซึ่งจดทะเบียนเป็นนิติบุคคล ณ {result.ContractorCompany} มีสำนักงานใหญ่อยู่เลขที่ {result.ContractorAddressNo} ถนน {result.ContractorStreet} ตำบล/แขวง {result.ContractorSubDistrict} อำเภอ/เขต {result.ContractorDistrict} จังหวัด {result.ContractorProvince} รหัสไปรษณีย์ {result.ContractorZipcode} </br>โดย {result.ContractorSignatoryName} ตำแหน่ง {result.ContractorSignatoryPosition} มีอำนาจลงนามผูกพันนิติบุคคลปรากฏตามหนังสือรับรองของสำนักงานทะเบียนหุ้นส่วนบริษัท ลงวันที่ {strContractorAuthDate} (และหนังสือมอบอำนาจเลขที่ {result.ContractorAuthNumber}) แนบท้ายสัญญานี้</p>"
                : $"<p class='t-16'>ในกรณีที่ผู้รับจ้างเป็นบุคคลธรรมดา กับ {result.ContractorName} อยู่บ้านเลขที่ {result.ContractorAddressNo} ถนน {result.ContractorStreet} ตำบล/แขวง {result.ContractorSubDistrict} อำเภอ/เขต {result.ContractorDistrict} จังหวัด {result.ContractorProvince} รหัสไปรษณีย์ {result.ContractorZipcode} ผู้ถือบัตรประจำตัวประชาชนเลขที่ {result.CitizenId} ซึ่งต่อไปในสัญญานี้เรียกว่า “ผู้ให้เช่า” อีกฝ่ายหนึ่ง</p>")}
   

<p class='tab3 t-16'>คู่สัญญาได้ตกลงกันมีข้อความดังต่อไปนี้</p>  
    <p class='tab3 t-16'><b>ข้อ ๑ ข้อตกลงเช่า</b></p>
    <p class='tab3 t-16'>ผู้เช่าตกลงเช่าและผู้ให้เช่าตกลงให้เช่าเครื่องถ่ายเอกสาร ยี่ห้อ {result.RentalCopierBrand} รุ่น {result.RentalCopierModel} หมายเลขเครื่อง {result.RentalCopierNumber} จำนวน {result.RentalCopierAmount} เครื่อง ซึ่งต่อไปในสัญญานี้เรียกว่า “เครื่องถ่ายเอกสารที่เช่า” 
</br>เพื่อใช้ในกิจการของผู้เช่าตามเอกสารแนบท้ายสัญญา</p>
    <p class='tab3 t-16'>การเช่าเครื่องถ่ายเอกสารที่เช่าตามวรรคหนึ่งมีกำหนดระยะเวลา {result.RentalYears} ปี {result.RentalMonths} เดือน 
</br>ตั้งแต่ {strStartDate} ถึง {strEndDate}</p>
    <p class='tab3 t-16'>ผู้ให้เช่ารับรองว่าเครื่องถ่ายเอกสารที่เช่าตามสัญญานี้เป็นเครื่องถ่ายเอกสารใหม่ที่ไม่เคย
</br>ใช้งานมาก่อน ผู้ให้เช่าได้ชำระภาษี อากร ค่าธรรมเนียมต่างๆ ครบถ้วนถูกต้องตามกฎหมายแล้ว ผู้ให้เช่า
</br>มีสิทธินำมาให้เช่าโดยปราศจากการรอนสิทธิ ทั้งรับรองว่าเครื่องถ่ายเอกสารที่เช่ามีคุณสมบัติ คุณภาพ 
</br>และคุณลักษณะไม่ต่ำกว่าที่กำหนดไว้ในเอกสารแนบท้ายสัญญาผนวก และผู้ให้เช่าได้ตรวจสอบแล้วว่า
</br>เครื่องถ่ายเอกสารที่เช่าตลอดจนอุปกรณ์ทั้งปวงปราศจากความชำรุดบกพร่อง</p>

    <p class='tab3 t-16'><b>ข้อ ๒ ค่าเช่าเครื่องถ่ายเอกสาร</b></p>
    <p class='tab3 t-16'>ผู้เช่าตกลงชำระค่าเช่าแก่ผู้ให้เช่าเป็นรายเดือนตามเดือนปฏิทินในอัตราค่าเช่าเดือนละ {result.RatePerUnit?.ToString("N2") ?? "0.00"}
 บาท ({strPerUnit}) 
</br>ต่อเครื่องถ่ายเอกสารหนึ่งเครื่อง รวมเป็นค่าเช่าทั้งสิ้นเดือนละ {result.RateTotal?.ToString("N2") ?? "0.00"}
 บาท ({strRateTotal}) โดยประเมิน
</br>จากจำนวนสำเนาเอกสารที่ถ่ายทั้งสิ้นเดือนละ {result.EstCopiesPerMonth} แผ่น</p>
    <p class='tab3 t-16'>หากเดือนใดจำนวนสำเนาเอกสารที่ผู้เช่าได้ถ่ายจากเครื่องถ่ายเอกสารที่เช่ามีจำนวนทั้งสิ้นไม่ถึง {result.IfNotCopiesAmount} แผ่น การชำระค่าเช่าในเดือนนั้นให้เปลี่ยนเป็นคิดคำนวณจากจำนวนสำเนาเอกสารที่ถ่ายในเดือนนั้นๆ ในอัตราสำเนาแผ่นละ {result.CopiesRate?.ToString("N2") ?? "0.00"}
 บาท</p>
    <p class='tab3 t-16'>ค่าเช่าตามวรรคหนึ่งและวรรคสองได้รวมภาษีมูลค่าเพิ่ม ค่าใช้จ่ายในการบำรุงรักษาและซ่อม
</br>แซม ค่าตรวจสภาพ ค่าอะไหล่ ค่าวัสดุสิ้นเปลือง (ยกเว้นค่ากระดาษถ่ายเอกสาร) ไว้ด้วยแล้ว</p>
    <p class='tab3 t-16'>ในการชำระค่าเช่า ผู้ให้เช่าต้องส่งใบแจ้งหนี้เรียกเก็บค่าเช่าเมื่อสิ้นเดือนแต่ละเดือน โดยผู้เช่า
</br>จะชำระค่าเช่าหลังจากที่ได้ตรวจสอบแล้วว่าถูกต้อง</p>
    <p class='tab3 t-16'>การจ่ายเงินตามเงื่อนไขแห่งสัญญานี้ ผู้เช่าจะโอนเงินเข้าบัญชีเงินฝากธนาคารของผู้ให้เช่า
</br>ชื่อธนาคาร {result.SaleBankName} สาขา {result.SaleBankBranch} ชื่อบัญชี {result.SaleBankAccountName} เลขที่บัญชี {result.SaleBankAccountNumber} ทั้งนี้ 
</br>ผู้ให้เช่าตกลงเป็นผู้รับภาระเงินค่าธรรมเนียม หรือค่าบริการอื่นใดเกี่ยวกับการโอน รวมทั้งค่าใช้จ่ายอื่นใด (ถ้ามี) ที่ธนาคารเรียกเก็บ และยินยอมให้มีการหักเงินดังกล่าวจากจำนวนเงินโอนในงวดนั้นๆ</p>
    <p class='tab3 t-16'>ข้อ ๓ เอกสารอันเป็นส่วนหนึ่งของสัญญา</p>
    <p class='tab3 t-16'>เอกสารแนบท้ายสัญญาดังต่อไปนี้ให้ถือเป็นส่วนหนึ่งของสัญญานี้</p>
{htmlDocAtt}
    <p class='tab3 t-16'>ความใดในเอกสารแนบท้ายสัญญาที่ขัดหรือแย้งกับข้อความในสัญญานี้ ให้ใช้ข้อความใน
</br>สัญญานี้บังคับ และในกรณีที่เอกสารแนบท้ายสัญญาขัดแย้งกันเอง ผู้ให้เช่าจะต้องปฏิบัติตามคำวินิจฉัย
</br>ของผู้เช่า คำวินิจฉัยของผู้เช่าให้ถือเป็นที่สุด และผู้ให้เช่าไม่มีสิทธิเรียกร้องค่าเช่า ค่าเสียหาย หรือค่าใช้จ่าย
</br>ใดๆ เพิ่มเติมจากผู้เช่าทั้งสิ้น
</p>
    <p class='tab3 t-16'><b>ข้อ ๔ การส่งมอบ</b></p>
    <p class='tab3 t-16'>ผู้ให้เช่าต้องส่งมอบและติดตั้งเครื่องถ่ายเอกสารที่เช่าตามสัญญานี้ ให้ถูกต้องครบถ้วนตาม
</br>สัญญานี้ ในลักษณะพร้อมใช้งานได้ตามที่กำหนด ณ {result.DeliveryLocation} ภายในวันที่ {CommonDAO.ToThaiDateStringCovert(result.DeliveryDate ?? DateTime.Now)} ซึ่งผู้ให้เช่าเป็นผู้จัดหาอุปกรณ์ประกอบพร้อมทั้งเครื่องมือ 
</br>ที่จำเป็นในการติดตั้งและใช้งาน โดยผู้ให้เช่าเป็นผู้ออกค่าใช้จ่ายเองทั้งสิ้น</p>
   

<p class='tab3 t-16'>ผู้ให้เช่าต้องแจ้งเวลาติดตั้งแล้วเสร็จพร้อมที่จะใช้งานและส่งมอบเครื่องได้เป็นหนังสือ 
</br>ต่อผู้เช่า ณ {result.DeliveryLocation} ในวันและเวลาทำการของผู้เช่าก่อนวันกำหนดส่งมอบตามวรรคหนึ่งไม่น้อยกว่า {result.DeliveryType} วันทำการของผู้เช่า</p>
    <p class='tab3 t-16'>ในการส่งมอบตามวรรคหนึ่ง ผู้ให้เช่าต้องส่งพนักงานมาดำเนินการทดสอบประสิทธิภาพและ
</br>แนะนำวิธีการใช้เครื่องให้คณะกรรมการตรวจรับพัสดุได้พิจารณาตามรายละเอียดคุณลักษณะเฉพาะที่ระบุไว้
</br>ในข้อ 1 และสำเนาที่ถ่ายจะต้องมีความชัดเจนสะอาด ไม่มีรอยหมึกเปื้อนตามส่วนต่างๆ โดยในการนี้ผู้ให้เช่า
</br>ไม่คิดค่าใช้จ่ายใดๆ จากผู้เช่าทั้งสิ้น</p>
   

<p class='tab3 t-16'><b>ข้อ ๕ การตรวจรับ</b></p>
    <p class='tab3 t-16'>เมื่อผู้เช่าได้ตรวจรับเครื่องถ่ายเอกสารที่ส่งมอบตามข้อ 4 และเห็นว่าถูกต้องครบถ้วนตาม
</br>สัญญานี้แล้ว ผู้เช่าจะออกหลักฐานการรับมอบเครื่องถ่ายเอกสารที่เช่าไว้เป็นหนังสือ เพื่อผู้ให้เช่านำมาใช้
</br>เป็นหลักฐานประกอบการขอรับเงินค่าเช่า</p>
    <p class='tab3 t-16'>ในการตรวจรับเครื่องถ่ายเอกสารที่ส่งมอบตามวรรคหนึ่ง ถ้าปรากฏว่าเครื่องถ่ายเอกสารซึ่งผู้
ให้เช่าส่งมอบไม่ถูกต้องครบถ้วนตามสัญญา หรือติดตั้งและส่งมอบถูกต้องครบถ้วนภายในกำหนดแต่ไม่สามารถ
ใช้งานได้อย่างครบถ้วนและมีประสิทธิภาพตามสัญญา ผู้เช่าทรงไว้ซึ่งสิทธิที่จะไม่รับเครื่องถ่ายเอกสารนั้นใน
</br>กรณีเช่นว่านี้ ผู้ให้เช่าต้องรีบนำเครื่องถ่ายเอกสารนั้นกลับคืนไปทันที และต้องนำเครื่องถ่ายเอกสารเครื่องใหม่
</br>ที่มีคุณสมบัติเดียวกัน หรือไม่ต่ำกว่าเครื่องถ่ายเอกสารที่กำหนดไว้ในสัญญานี้ มาส่งมอบให้ใหม่ ภายใน {result.TotalDay} วัน
</br>ด้วยค่าใช้จ่ายของผู้ให้เช่าเองทั้งสิ้น และระยะเวลาที่เสียไปเพราะเหตุดังกล่าวผู้ให้เช่าจะนำมาอ้างเป็น
</br>เหตุของดหรือลดค่าปรับหรือขยายเวลาส่งมอบไม่ได้</p>
    <p class='tab3 t-16'>หากผู้ให้เช่าไม่นำเครื่องถ่ายเอกสารที่ส่งมอบไม่ถูกต้องกลับคืนไปในทันทีดังกล่าวในวรรคสอง
</br>และเกิดความเสียหายแก่เครื่องถ่ายเอกสารนั้น ผู้เช่าไม่ต้องรับผิดชอบในความเสียหายดังกล่าว</p>
    <p class='tab3 t-16'>ในกรณีที่ผู้ให้เช่าส่งมอบเครื่องถ่ายเอกสารที่เช่าถูกต้องแต่ไม่ครบจำนวน หรือส่งมอบครบ
</br>จำนวนแต่ไม่ถูกต้องทั้งหมด ผู้เช่ามีสิทธิจะรับมอบเฉพาะส่วนที่ถูกต้อง โดยออกหลักฐานการรับมอบเฉพาะ
</br>ส่วนนั้นก็ได้ ในกรณีเช่นนี้ผู้เช่าจะชำระค่าเช่าเฉพาะเครื่องถ่ายเอกสารที่เช่าที่รับมอบไว้</p>
    

<p class='tab3 t-16'><b>ข้อ ๖ การงดหรือลดค่าปรับ หรือขยายเวลาในการปฏิบัติตามสัญญา</b></p>
    <p class='tab3 t-16'>ในกรณีที่ผู้ให้เช่าไม่สามารถส่งมอบเครื่องถ่ายเอกสารที่เช่าให้แก่ผู้เช่าได้โดยครบถ้วนถูกต้อง
</br>ภายในกำหนดเวลาตามสัญญา หรือถ้าผู้ให้เช่าไม่ดำเนินการหรือไม่สามารถซ่อมแซมแก้ไขเครื่องถ่ายเอกสาร
</br>ที่เช่าภายในระยะเวลาตามข้อ ๘.๒ และผู้ให้เช่าไม่จัดหาเครื่องถ่ายเอกสารให้ผู้เช่าใช้แทนตามข้อ ๘.๓ อันเนื่อง
</br>มาจากเหตุสุดวิสัย หรือเหตุใดๆ อันเนื่องมาจากความผิดหรือความบกพร่องของฝ่ายผู้เช่าหรือจากพฤติการณ์
</br>อันหนึ่งอันใดที่ผู้ให้เช่าไม่ต้องรับผิดตามกฎหมาย หรือเหตุอื่นตามที่กำหนดในกฎกระทรวง ซึ่งออกตามความ
</br>ในกฎหมายว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ ผู้ให้เช่ามีสิทธิของดหรือลดค่าปรับหรือขยาย
</br>กำหนดเวลาทำการตามสัญญาดังกล่าวโดยจะต้องแจ้งเหตุหรือพฤติการณ์ดังกล่าวพร้อมหลักฐานเป็นหนังสือ
</br>ให้ผู้เช่าทราบภายใน ๑๕ วันนับถัดจากวันที่เหตุนั้นสิ้นสุดลงหรือตามที่กำหนดในกฎกระทรวงดังกล่าว
</br>แล้วแต่กรณี</p>
    <p class='tab3 t-16'>ถ้าผู้ให้เช่าไม่ปฏิบัติให้เป็นไปตามความในวรรคหนึ่ง ให้ถือว่าผู้ให้เช่าได้สละสิทธิเรียกร้อง
</br>ในการที่จะของดหรือลดค่าปรับหรือขยายเวลาทำการตามสัญญาโดยไม่มีเงื่อนไขใดๆ ทั้งสิ้น เว้นแต่กรณี
</br>เหตุเกิดจากความผิดหรือความบกพร่องของฝ่ายผู้เช่าซึ่งมีหลักฐานชัดแจ้ง หรือผู้เช่าทราบดีอยู่แล้วตั้งแต่ต้น</p>
    <p class='tab3 t-16'>การงดหรือลดค่าปรับหรือขยายกำหนดเวลาทำการตามสัญญาตามวรรคหนึ่ง อยู่ในดุลพินิจ
</br>ของผู้เช่าที่จะพิจารณาตามที่เห็นสมควร</p>
   
<p class='tab3 t-16'><b>ข้อ ๗ การบำรุงรักษาตรวจสภาพและซ่อมแซมเครื่องถ่ายเอกสารที่เช่า</b></p>
    <p class='tab3 t-16'>ผู้ให้เช่ามีหน้าที่บำรุงรักษาเครื่องถ่ายเอกสารที่เช่าให้อยู่ในสภาพใช้งานได้ดีอยู่เสมอด้วยค่าใช้
</br>จ่ายของผู้ให้เช่า โดยต้องจัดหาช่างผู้มีความรู้ ความชำนาญ และฝีมือดีมาตรวจสอบ บำรุงรักษาและซ่อมแซม
</br>แก้ไขเครื่องถ่ายเอกสารที่เช่าตลอดอายุสัญญาเช่านี้ อย่างน้อยเดือนละ {result.MaintenancePermonth} ครั้ง โดยให้มีระยะเวลาห่างกันไม่
</br>น้อยกว่า {result.MaintenanceInterval} วัน</p>
    <p class='tab3 t-16'>สิ่งของที่ใช้สิ้นเปลืองทุกชนิดรวมทั้งอะไหล่ ยกเว้นกระดาษสำหรับถ่ายเอกสาร ผู้ให้เช่า

</br>จะเป็นผู้จัดส่งให้โดยไม่คิดมูลค่า โดยที่ผู้ให้เช่าจะจัดให้มีไว้ในความครอบครองของผู้เช่าให้เพียงพออยู่เสมอ
</br>อุปกรณ์สิ้นเปลืองดังกล่าว เช่น ลูกโม่ถ่ายภาพ ผงหมึก ผงประจุภาพ หมึกพิมพ์ วัสดุที่ใช้ทำความสะอาด
</br>ถุงกรอง แปรง น้ำมันหล่อลื่น และอุปกรณ์อื่นๆ ที่จำเป็นเพื่อให้เครื่องถ่ายเอกสารใช้งานได้ตามปกติตลอดเวลา</p>
  
<p class='tab3 t-16'><b>ข้อ ๘ หน้าที่ของผู้ให้เช่า</b></p>
    <p class='tab3 t-16'>๘.๑ ผู้ให้เช่ามีหน้าที่ฝึกอบรมวิธีใช้เครื่องถ่ายเอกสารที่เช่าให้แก่เจ้าหน้าที่ของผู้เช่า จน
</br>สามารถใช้งานเครื่องถ่ายเอกสารได้ และผู้ให้เช่าตกลงจะฝึกอบรมวิธีการใช้เครื่องถ่ายเอกสารที่เช่าให้แก่
</br>เจ้าหน้าที่ของผู้เช่าทุกครั้ง หากผู้เช่าร้องขอโดยเหตุที่มีการเปลี่ยนแปลงโยกย้ายเจ้าหน้าที่ของผู้เช่าและ
</br>เจ้าหน้าที่คนนั้นยังไม่เคยได้รับการฝึกอบรมมาก่อนโดยผู้ให้เช่าเป็นผู้รับผิดชอบค่าใช้จ่ายในการฝึกอบรมทั้งสิ้น</p>
    <p class='tab3 t-16'>๘.๒ ในกรณีเครื่องถ่ายเอกสารที่เช่าชำรุดบกพร่องหรือขัดข้องใช้งานไม่ได้ตามปกติ ผู้ให้
</br>เช่าจะต้องจัดให้ช่างที่มีความรู้ความชำนาญและฝีมือดีมาจัดการซ่อมแซมแก้ไขให้อยู่ในสภาพใช้งานได้ดี
</br>ตามปกติ โดยผู้ให้เช่าจะต้องเริ่มจัดการซ่อมแซมแก้ไขในทันทีที่ได้รับแจ้งจากผู้เช่าหรือผู้ที่ได้รับมอบหมาย
</br>จากผู้เช่าแล้ว และให้แล้วเสร็จใช้งานได้ดีดังเดิมอย่างช้าต้องไม่เกิน {result.CopierFixDays} ชั่วโมง ตั้งแต่เวลาที่ได้รับแจ้ง</p>
    <p class='tab3 t-16'>๘.๓ ในกรณีที่เครื่องถ่ายเอกสารที่เช่ามีความชำรุดบกพร่องหรือขัดข้องใช้งานไม่ได้ตามปกติ 
</br>และการซ่อมแซมต้องใช้เวลาเกินกว่า {result.ReplaceFixDays} ชั่วโมง ตามที่กำหนดในข้อ ๘.๒ หรือไม่อาจซ่อมแซมแก้ไขให้ดีได้
</br>ดังเดิม ผู้ให้เช่าต้องจัดหาเครื่องถ่ายเอกสารที่มีคุณสมบัติ คุณภาพ ความสามารถ และประสิทธิภาพในการ
</br>ใช้งานไม่ต่ำกว่าของเครื่องเดิมมาให้ผู้เช่าใช้แทนทันที</p>
    <p class='tab3 t-16'><b>ข้อ ๙ ค่าปรับกรณีความชำรุดบกพร่องของเครื่องถ่ายเอกสาร</b></p>
    <p class='tab3 t-16'>ถ้าผู้ให้เช่าไม่ดำเนินการหรือไม่สามารถซ่อมแซมแก้ไขเครื่องถ่ายเอกสารที่เช่าภายในระยะ
</br>เวลาตามข้อ ๘.๒ และผู้ให้เช่าไม่จัดหาเครื่องถ่ายเอกสารให้ผู้เช่าใช้แทนตามข้อ ๘.๓ ผู้ให้เช่ายินยอมให้
</br>ผู้เช่าปรับเป็นรายวัน ในอัตราวันละ {result.FinePerDays?.ToString("N2") ?? "0.00"}
 บาท ({strFinePerDays}) ต่อเครื่อง ตั้งแต่พ้นกำหนดระยะเวลา
</br>ตามข้อ ๘.๒ จนถึงวันที่ผู้ให้เช่าซ่อมแซมแก้ไขให้อยู่ในสภาพใช้งานได้ดีตามปกติ หรือผู้ให้เช่าจัดหาเครื่องถ่าย
</br>เอกสารมาให้ผู้เช่าใช้งานแทน หรือจนกว่าผู้เช่าจะใช้สิทธิบอกเลิกสัญญา ทั้งนี้ ผู้เช่าไม่ต้องจ่ายค่าเช่าใน
</br>ระหว่างเวลาที่ผู้เช่าไม่สามารถใช้เครื่องถ่ายเอกสารที่เช่าตามสัญญานี้ โดยยินยอมให้ผู้เช่าหักค่าปรับดังกล่าว
</br>ออกจากค่าเช่าตามข้อ ๒ หรือบังคับเอาจากหลักประกันตามข้อ ๑๐ ก็ได้</p>

    <p class='tab3 t-16'><b>ข้อ ๑๐ หลักประกันการปฏิบัติตามสัญญา</b></p>
    <p class='tab3 t-16'>ในขณะทำสัญญานี้ผู้ให้เช่าได้นำหลักประกันเป็น {result.GuaranteeType} เป็นจำนวนเงิน {result.GuaranteeAmount?.ToString("N2") ?? "0.00"}
 บาท ({strGuaranteeAmount}) ซึ่งเท่ากับร้อยละ {result.GuaranteePercent} ของค่าเช่าทั้งหมดตามสัญญา มามอบ
</br>ให้แก่ผู้เช่าเพื่อเป็นหลักประกันการปฏิบัติตามสัญญานี้</p>
    <p class='tab3 t-16'>กรณีผู้ให้เช่าใช้หนังสือค้ำประกันมาเป็นหลักประกันการปฏิบัติตามสัญญา หนังสือค้ำประกัน
</br>ดังกล่าวจะต้องออกโดยธนาคารที่ประกอบกิจการในประเทศไทย หรือโดยบริษัทเงินทุนหรือบริษัทเงินทุน
</br>หลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศ
</br>ของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบตามแบบ
</br>ที่คณะกรรมการนโยบายการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐกำหนด หรืออาจเป็นหนังสือค้ำประกัน
</br>อิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้ และจะต้องมีอายุการค้ำประกันตลอดไปจนกว่าผู้ให้เช่า
</br>พ้นข้อผูกพันตามสัญญานี้</p>
    <p class='tab3 t-16'>หลักประกันที่ผู้ให้เช่านำมามอบให้ตามวรรคหนึ่ง จะต้องมีอายุครอบคลุมความรับผิดทั้งปวง
</br>ของผู้ให้เช่าตลอดอายุสัญญา ถ้าหลักประกันที่ผู้ให้เช่านำมามอบให้ดังกล่าวลดลงหรือเสื่อมค่าลง หรือมีอายุ
</br>ไม่ครอบคลุมถึงความรับผิดของผู้ให้เช่าตลอดอายุสัญญา ไม่ว่าด้วยเหตุใดๆ ก็ตาม รวมถึงกรณีผู้ให้เช่าส่งมอบ
</br>และติดตั้งเครื่องถ่ายเอกสารล่าช้าเป็นเหตุให้ระยะเวลาการเช่าตามสัญญาเปลี่ยนแปลงไป ผู้ให้เช่าต้องหาหลัก
</br>ประกันใหม่หรือหลักประกันเพิ่มเติมให้มีจำนวนครบถ้วนตามวรรคหนึ่งมามอบให้แก่ผู้เช่าภายใน {result.NewGuaranteeDays} วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้เช่า</p>
    <p class='tab3 t-16'>หลักประกันที่ผู้ให้เช่านำมามอบไว้ตามข้อนี้ ผู้เช่าจะคืนให้แก่ผู้ให้เช่าโดยไม่มีดอกเบี้ยเมื่อ
</br>ผู้ให้เช่าพ้นจากข้อผูกพันและความรับผิดทั้งปวงตามสัญญานี้แล้ว</p>
    <p class='tab3 t-16'><b>ข้อ ๑๑ การบอกเลิกสัญญา</b></p>
    <p class='tab3 t-16'>เมื่อครบกำหนดส่งมอบเครื่องถ่ายเอกสารที่เช่าตามสัญญาแล้ว ถ้าผู้ให้เช่าไม่ส่งมอบเครื่อง
</br>ถ่ายเอกสารที่เช่า หรือส่งมอบแต่เพียงบางส่วนให้แก่ผู้เช่า หรือส่งมอบเครื่องถ่ายเอกสารที่เช่าไม่ตรงตาม
</br>สัญญาหรือมีคุณลักษณะเฉพาะไม่ถูกต้องตามสัญญา หรือส่งมอบแล้วเสร็จภายในกำหนดแต่ไม่สามารถ
</br>ใช้งานได้อย่างมีประสิทธิภาพหรือใช้งานได้ไม่ครบถ้วนตามสัญญา หรือผู้ให้เช่าไม่ปฏิบัติตามสัญญาข้อใด
</br>ข้อหนึ่งผู้เช่ามีสิทธิบอกเลิกสัญญาทั้งหมดหรือแต่บางส่วนได้ การใช้สิทธิบอกเลิกสัญญานั้นไม่กระทบสิทธิ
</br>ของผู้เช่าที่จะเรียกร้องค่าเสียหายจากผู้ให้เช่า</p>
    <p class='tab3 t-16'>ในกรณีที่ผู้เช่าใช้สิทธิบอกเลิกสัญญา ผู้เช่ามีสิทธิริบหรือบังคับจากหลักประกัน ตามข้อ ๑๐ 
</br>เป็นจำนวนเงินทั้งหมดหรือแต่บางส่วนก็ได้แล้วแต่ผู้เช่าจะเห็นสมควร และถ้าผู้เช่าต้องเช่าเครื่องถ่ายเอกสาร
</br>จากบุคคลอื่นทั้งหมดหรือแต่บางส่วนภายในกำหนด {result.TeminationReplaceDays} เดือน นับถัดจากวันบอกเลิกสัญญา ผู้ให้เช่ายอมรับผิด
</br>ชดใช้ค่าเช่าที่เพิ่มขึ้นจากค่าเช่าที่กำหนดไว้ในสัญญานี้ รวมทั้งค่าใช้จ่ายใดๆ ที่ผู้เช่าต้องใช้จ่ายในการจัดหา
</br>ผู้ให้เช่าเครื่องถ่ายเอกสารที่เช่ารายใหม่ดังกล่าวด้วย</p>
    <p class='tab3 t-16'>ในกรณีมีความจำเป็น ผู้เช่ามีสิทธิที่จะบอกเลิกสัญญาเช่านี้ก่อนครบกำหนดระยะเวลาการ
</br>เช่าได้ โดยแจ้งเป็นหนังสือให้ผู้ให้เช่าทราบล่วงหน้าไม่น้อยกว่า ๓๐ วัน โดยผู้ให้เช่าจะไม่มีสิทธิเรียกร้อง
</br>ค่าเสียหายใดๆ จากผู้เช่า</p>
    
<p class='tab3 t-16'><b>ข้อ ๑๒ ค่าปรับกรณีส่งมอบล่าช้า</b></p>
    <p class='tab3 t-16'>ในกรณีที่ผู้ให้เช่าส่งมอบเครื่องถ่ายเอกสารที่เช่าล่วงเลยกำหนดส่งมอบตามข้อ ๔ และผู้เช่า
</br>มิได้ใช้สิทธิบอกเลิกสัญญาตามข้อ ๑๑ วรรคหนึ่ง ผู้ให้เช่าจะต้องชำระค่าปรับให้ผู้เช่าเป็นรายวัน สำหรับ
</br>เครื่องถ่ายเอกสารที่ยังไม่ได้ส่งมอบตามสัญญา ในอัตราวันละ {result.LateFinePerDays?.ToString("N2") ?? "0.00"}
 บาท ({strLate}) ต่อเครื่อง นับถัดจาก
</br>วันที่ครบกำหนดส่งมอบตามสัญญาจนถึงวันที่ผู้ให้เช่าได้นำเครื่องถ่ายเอกสารที่เช่ามาส่งมอบให้แก่ผู้เช่า
</br>จนถูกต้องครบถ้วน</p>
    <p class='tab3 t-16'>ในระหว่างที่ผู้เช่ายังมิได้ใช้สิทธิบอกเลิกสัญญานั้น หากผู้เช่าเห็นว่าผู้ให้เช่าไม่อาจปฏิบัติตาม
</br>สัญญาต่อไปได้ ผู้เช่าจะใช้สิทธิบอกเลิกสัญญา และบังคับจากหลักประกันการปฏิบัติตามสัญญาตามข้อ ๑๐ 
</br>กับเรียกร้องให้ชดใช้ค่าเช่าที่เพิ่มขึ้นตามที่กำหนดไว้ในข้อ ๑๑ วรรคสองก็ได้ และถ้าผู้เช่าได้แจ้งข้อเรียกร้อง
</br>ให้ชำระค่าปรับไปยังผู้ให้เช่าเมื่อครบกำหนดส่งมอบดังกล่าวแล้ว ผู้เช่ามีสิทธิที่จะปรับผู้ให้เช่าจนถึงวันบอก
</br>เลิกสัญญาได้อีกด้วย</p>

    <p class='tab3 t-16'><b>ข้อ ๑๓ การบังคับค่าปรับ ค่าเสียหาย และค่าใช้จ่าย</b></p>
    <p class='tab3 t-16'>ในกรณีที่ผู้ให้เช่าไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่งด้วยเหตุใดๆ ก็ตาม จนเป็นเหตุให้เกิดค่า
</br>ปรับ ค่าเสียหาย หรือค่าใช้จ่ายแก่ผู้เช่า ผู้ให้เช่าต้องชดใช้ค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายดังกล่าวให้แก่ผู้เช่า
</br>โดยสิ้นเชิงภายในกำหนด {result.EnforcementOfFineDays} วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้เช่า หากผู้ให้เช่าไม่ชดใช้ให้ถูก
</br>ต้องครบถ้วนภายในระยะเวลาดังกล่าวให้ผู้เช่ามีสิทธิที่จะหักเอาจากค่าเช่าที่ต้องชำระหรือบังคับจากหลัก
</br>ประกันการปฏิบัติตามสัญญาได้ทันที</p>
    <p class='tab3 t-16'>หากค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายที่บังคับจากค่าเช่าที่ต้องชำระ หรือหลักประกันการ
</br>ปฏิบัติตามสัญญาแล้วยังไม่เพียงพอ ผู้ให้เช่ายินยอมชำระส่วนที่เหลือที่ยังขาดอยู่จนครบถ้วนตามจำนวน
</br>ค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายนั้น ภายในกำหนด {result.OutstandingPeriodDays} วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือ
</br>จากผู้เช่า</p>
    <p class='tab3 t-16'><b>ข้อ ๑๔ การโอนสิทธิของผู้ให้เช่า</b></p>
    <p class='tab3 t-16'>ในระหว่างอายุสัญญาเช่า ห้ามผู้ให้เช่าโอนสิทธิหน้าที่ตามสัญญาหรือกรรมสิทธิ์ในเครื่องถ่าย
</br>เอกสารที่เช่าแก่บุคคลอื่น เว้นแต่จะได้รับความยินยอมเป็นหนังสือจากผู้เช่าก่อน</p>
    <p class='tab3 t-16'><b>ข้อ ๑๕ การนำเครื่องถ่ายเอกสารที่เช่ากลับคืนเมื่อสัญญาสิ้นสุดลง</b></p>
    <p class='tab3 t-16'>เมื่อสัญญาสิ้นสุดลงไม่ว่าจะเป็นการบอกเลิกสัญญาหรือครบกำหนดเวลาตามสัญญา ผู้ให้เช่า
</br>ต้องนำเครื่องถ่ายเอกสารที่เช่ากลับคืนไปภายใน {result.CopierSendBackDays} วัน โดยผู้ให้เช่าเป็นผู้เสียค่าใช้จ่ายเองทั้งสิ้น</p>
    <p class='tab3 t-16'>ถ้าผู้ให้เช่าไม่นำเครื่องถ่ายเอกสารที่เช่ากลับคืนไปภายในกำหนดเวลาตามวรรคหนึ่ง ผู้เช่าไม่
</br>ต้องรับผิดชอบในความเสียหายใดๆ ทั้งสิ้นที่เกิดแก่เครื่องถ่ายเอกสารที่เช่าอันมิใช่ความผิดของผู้เช่า</p>
    <p class='tab3 t-16'><b>ข้อ ๑๖ ข้อจำกัดความรับผิดของผู้เช่า</b></p>
    <p class='tab3 t-16'>ผู้เช่าไม่ต้องรับผิดในความเสียหายหรือสูญหายเมื่อเกิดอัคคีภัยหรือภัยพิบัติใดๆ หรือการ
</br>โจรกรรมเครื่องถ่ายเอกสารที่เช่าตลอดจนการสูญหายหรือความเสียหายใดๆ ที่เกิดขึ้นแก่เครื่องถ่ายเอกสาร
</br>ที่เช่าอันไม่ใช่เกิดจากความผิดของผู้เช่าตลอดระยะเวลาที่เครื่องถ่ายเอกสารอยู่ในความครอบครองของผู้เช่า</p>
    <p class='tab3 t-16'>สัญญานี้ทำขึ้นสองฉบับมีข้อความถูกต้องตรงกัน คู่สัญญาได้อ่านและเข้าใจข้อความโดย
</br>ละเอียดตลอดแล้ว จึงได้ลงลายมือชื่อ พร้อมประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยาน และคู่สัญญาต่าง
</br>ยึดถือไว้คนละหนึ่งฉบับ</p>


</br>
</br>
{signatoryHtml}


    <div style='page-break-before: always;'></div>
     <p class='text-center t-22' style='font-weight:bold;'>วิธีปฏิบัติเกี่ยวกับสัญญาเช่าเครื่องถ่ายเอกสาร</p>
     <p class='tab2 t-16'>(๑) ให้ระบุเลขที่สัญญาในปีงบประมาณหนึ่งๆ ตามลำดับ</p>
     <p class='tab2 t-16'>(๒) ให้ระบุชื่อของหน่วยงานของรัฐที่เป็นนิติบุคคล เช่น กรม ก. หรือรัฐวิสาหกิจ ข. เป็นต้น</p>
     <p class='tab2 t-16'>(๓) ให้ระบุชื่อและตำแหน่งของหัวหน้าหน่วยงานของรัฐที่เป็นนิติบุคคลนั้น หรือผู้ที่ได้รับมอบอำนาจ เช่น นาย ก. อธิบดีกรม หรือ นาย ข. ผู้ได้รับมอบอำนาจจากอธิบดีกรม</p>
     <p class='tab2 t-16'>(๔) ให้ระบุชื่อผู้ให้เช่า</p>
    <p class='tab3 t-16'>ก. กรณีนิติบุคคล เช่น ห้างหุ้นส่วนสามัญจดทะเบียน ห้างหุ้นส่วนจำกัด บริษัทจำกัด</p>
    <p class='tab3 t-16'>ข. กรณีบุคคลธรรมดา ให้ระบุชื่อและที่อยู่</p>
     <p class='tab2 t-16'>(๕) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
     <p class='tab2 t-16'>(๖) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
     <p class='tab2 t-16'>(๗) หน่วยงานของรัฐอาจกำหนดเงื่อนไขการจ่ายค่าเช่าให้แตกต่างไปจากแบบสัญญาที่กำหนดได้ตามความเหมาะสมและจำเป็นและไม่ทำให้หน่วยงานของรัฐเสียเปรียบ หากหน่วยงานของรัฐเห็นว่าจะมีปัญหาในทางเสียเปรียบหรือไม่รัดกุมพอ ก็ให้ส่งร่างสัญญานั้นไปให้สำนักงานอัยการสูงสุดพิจารณาให้ความเห็นชอบก่อน</p>
     <p class='tab2 t-16'>(๘) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
     <p class='tab2 t-16'>(๙) ชื่อสถานที่หน่วยงานของรัฐ</p>
     <p class='tab2 t-16'>(๑๐) ให้พิจารณาถึงความจำเป็นและเหมาะสมของการใช้งาน</p>
     <p class='tab2 t-16'>(๑๑) อัตราค่าปรับตามสัญญาข้อ ๙ ให้กำหนดเป็นรายวันในอัตราระหว่างร้อยละ ๐.๐๑ – ๐.๒๐ ตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ ข้อ ๑๖๒ ส่วนกรณีจะปรับร้อยละเท่าใด ให้อยู่ในดุลยพินิจของหน่วยงานของรัฐผู้เช่าที่จะพิจารณา โดยคำนึงถึงราคาและลักษณะของพัสดุที่เช่า ซึ่งอาจมีผลกระทบต่อการที่ผู้ให้เช่าจะหลีกเลี่ยงไม่ปฏิบัติตามสัญญา แต่ทั้งนี้การที่จะกำหนดค่าปรับเป็นร้อยละเท่าใดจะต้องกำหนดไว้ในเอกสารเชิญชวนด้วย</p>
     <p class='tab2 t-16'>(๑๒) “หลักประกัน” หมายถึง หลักประกันที่ผู้ให้เช่านำมามอบไว้แก่หน่วยงานของรัฐ เมื่อลงนามในสัญญา เพื่อเป็นการประกันความเสียหายที่อาจจะเกิดขึ้นจากการปฏิบัติตามสัญญา ดังนี้</p>
    <p class='tab3 t-16'>(๑) เงินสด</p>
    <p class='tab3 t-16'>(๒) เช็คหรือดราฟท์ ที่ธนาคารเซ็นสั่งจ่าย ซึ่งเป็นเช็คหรือดราฟท์ลงวันที่ที่ใช้เช็คหรือดราฟท์นั้นชำระต่อเจ้าหน้าที่ หรือก่อนวันนั้นไม่เกิน ๓ วันทำการ</p>
    <p class='tab3 t-16'>(๓) หนังสือคํ้าประกันของธนาคารภายในประเทศตามตัวอย่างที่คณะกรรมการนโยบายกำหนด โดยอาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้</p>
    <p class='tab3 t-16'>(๔) หนังสือค้ำประกันของบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบ โดยอนุโลมให้ใช้ตามตัวอย่างหนังสือค้ำประกันของธนาคารที่คณะกรรมการนโยบายกำหนด</p>
    <p class='tab3 t-16'>(๕) พันธบัตรรัฐบาลไทย</p>
     <p class='tab2 t-16'>(๑๓) ให้กำหนดจำนวนเงินหลักประกันการปฏิบัติตามสัญญาตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ ข้อ ๑๖๘</p>
     <p class='tab2 t-16'>(๑๔) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง</p>
     <p class='tab2 t-16'>(๑๕) กำหนดระยะเวลาตามความเหมาะสม เช่น 3 เดือน</p>
    <p class='tab3 t-16'>(๑๖) อัตราค่าปรับตามสัญญาข้อ 12 ให้กำหนดเป็นรายวันในอัตราระหว่างร้อยละ ๐.๐๑ – ๐.๒๐ ตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. ๒๕๖๐ ข้อ ๑๖๒ ส่วนกรณีจะปรับร้อยละเท่าใด ให้อยู่ในดุลยพินิจของหน่วยงานของรัฐผู้เช่าที่จะพิจารณา โดยคำนึงถึงราคาและลักษณะของพัสดุที่เช่า ซึ่งอาจมีผลกระทบต่อการที่ผู้ให้เช่าจะหลีกเลี่ยงไม่ปฏิบัติตามสัญญา แต่ทั้งนี้การที่จะกำหนดค่าปรับเป็นร้อยละเท่าใดจะต้องกำหนดไว้ในเอกสารเชิญชวนด้วย</p>
</body>
</html>
";

        //    var doc = new HtmlToPdfDocument()
        //    {
        //        GlobalSettings = {
        //    PaperSize = PaperKind.A4,
        //    Orientation = Orientation.Portrait,
        //    Margins = new MarginSettings { Top = 20, Bottom = 20, Left = 20, Right = 20 }
        //},
        //        Objects = {
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
        //    };

        //    var pdfBytes = _pdfConverter.Convert(doc);
            return html;
        } catch (Exception ex) { throw new Exception("Error generating Word document", ex); }
       
    }

}
