using BatchAndReport.DAO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using System.Diagnostics.Contracts;
using System.Threading.Tasks;
namespace BatchAndReport.Pages.Report
{
    public class ExportModel : PageModel
    {
        private readonly SmeDAO _smeDao;
        private readonly WordEContract_AllowanceService _AllowanceService;
        private readonly WordEContract_LoanPrinterService _wordEContract_LoanPrinterService;
        private readonly WordEContract_ContactToDoThingService _ContactToDoThingService;
        private readonly WordEContract_HireEmployee _HireEmployee;
        private readonly WordEContract_BorrowMoneyService _BorrowMoneyService;
        private readonly WordEContract_MaintenanceComputerService _maintenanceComputerService;
        private readonly WordEContract_LoanComputerService _LoanComputerService;
        private readonly WordEContract_BuyAgreeProgram _BuyAgreeProgram;
        private readonly WordEContract_BuyOrSellComputerService _BuyOrSellComputerService;
        private readonly WordEContract_BuyOrSellService _BuyOrSellService;
        private readonly WordEContract_DataSecretService _DataSecretService;
        private readonly WordEContract_MemorandumService _MemorandumService;
        private readonly WordEContract_PersernalProcessService _PersernalProcessService;
        private readonly WordEContract_SupportSMEsService _SupportSMEsService;
        private readonly WordEContract_JointOperationService _JointOperationService;
        private readonly WordEContract_ControlDataService _ControlDataService;
        private readonly WordEContract_DataPersonalService _DataPersonalService;
        private readonly WordEContract_ConsultantService _ConsultantService;
        private readonly WordEContract_Test_HeaderLOGOService _Test_HeaderLOGOService;
        public ExportModel(SmeDAO smeDao, WordEContract_AllowanceService allowanceService
            , WordEContract_LoanPrinterService wordEContract_LoanPrinterService
            , WordEContract_ContactToDoThingService ContactToDoThingService
            , WordEContract_HireEmployee HireEmployee
            , WordEContract_BorrowMoneyService BorrowMoneyService
            , WordEContract_MaintenanceComputerService maintenanceComputerService
            , WordEContract_LoanComputerService LoanComputerService
            , WordEContract_BuyAgreeProgram BuyAgreeProgram
            , WordEContract_BuyOrSellComputerService BuyOrSellComputerService
            , WordEContract_BuyOrSellService BuyOrSellService
            , WordEContract_DataSecretService DataSecretService

            , WordEContract_MemorandumService MemorandumService
            , WordEContract_PersernalProcessService PersernalProcessService
            , WordEContract_SupportSMEsService SupportSMEsService
            , WordEContract_JointOperationService JointOperationService
            , WordEContract_ControlDataService ControlDataService
            , WordEContract_DataPersonalService DataPersonalService
            , WordEContract_ConsultantService ConsultantService
            , WordEContract_Test_HeaderLOGOService Test_HeaderLOGOService

            )
        {
            _smeDao = smeDao;
            _AllowanceService = allowanceService;
            this._wordEContract_LoanPrinterService = wordEContract_LoanPrinterService;
            _ContactToDoThingService = ContactToDoThingService;
            _HireEmployee = HireEmployee;
            _BorrowMoneyService = BorrowMoneyService;
            _maintenanceComputerService = maintenanceComputerService;
            this._LoanComputerService = LoanComputerService;
            this._BuyAgreeProgram = BuyAgreeProgram;
             _BuyOrSellComputerService = BuyOrSellComputerService;
            this._BuyOrSellService = BuyOrSellService;
            this._DataSecretService = DataSecretService;
            this._MemorandumService = MemorandumService;
            this._PersernalProcessService = PersernalProcessService;
            this._SupportSMEsService = SupportSMEsService;
            this._JointOperationService = JointOperationService;
            this._ControlDataService = ControlDataService;
            this._DataPersonalService = DataPersonalService;
            this._ConsultantService = ConsultantService;
            this._Test_HeaderLOGOService = Test_HeaderLOGOService;
        }
        public IActionResult OnGetPdf()
        {
            var wordDAO = new WordToPDFDAO(); // Create an instance of WordDAO
            var Resultpdf = wordDAO.OnGetPdfWithInterop(); // Call the method on the instance
            return Resultpdf; // Return an empty result since the PDF is handled in WordDAO
        }

        public IActionResult OnGetExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Products");

            // Header row
            ws.Cells["A1"].Value = "ProductID";
            ws.Cells["B1"].Value = "ProductCode";
            ws.Cells["C1"].Value = "ProductName";
            ws.Cells["D1"].Value = "Category";
            ws.Cells["E1"].Value = "Price";
            ws.Cells["F1"].Value = "StockQty";
            ws.Cells["G1"].Value = "Unit";
            ws.Cells["H1"].Value = "IsActive";
            ws.Cells["I1"].Value = "CreatedDate";

            // Sample data
            var products = new[]
            {
                    new { ProductID = 1, ProductCode = "P001", ProductName = "Apple iPhone 15", Category = "Mobile", Price = 35900, StockQty = 25, Unit = "pcs", IsActive = true, CreatedDate = new DateTime(2025, 1, 15, 10, 30, 0) },
                    new { ProductID = 2, ProductCode = "P002", ProductName = "Samsung Galaxy S24", Category = "Mobile", Price = 29900, StockQty = 40, Unit = "pcs", IsActive = true, CreatedDate = new DateTime(2025, 1, 18, 11, 0, 0) },
                    new { ProductID = 3, ProductCode = "P003", ProductName = "Dell XPS 13", Category = "Laptop", Price = 49900, StockQty = 15, Unit = "pcs", IsActive = true, CreatedDate = new DateTime(2025, 1, 20, 9, 45, 0) },
                    new { ProductID = 4, ProductCode = "P004", ProductName = "Logitech Mouse M590", Category = "Accessories", Price = 850, StockQty = 100, Unit = "pcs", IsActive = true, CreatedDate = new DateTime(2025, 2, 1, 8, 20, 0) },
                    new { ProductID = 5, ProductCode = "P005", ProductName = "HP Ink 678", Category = "Printer Ink", Price = 390, StockQty = 200, Unit = "pcs", IsActive = false, CreatedDate = new DateTime(2025, 2, 10, 12, 10, 0) }
                };

            int row = 2;
            foreach (var p in products)
            {
                ws.Cells[row, 1].Value = p.ProductID;
                ws.Cells[row, 2].Value = p.ProductCode;
                ws.Cells[row, 3].Value = p.ProductName;
                ws.Cells[row, 4].Value = p.Category;
                ws.Cells[row, 5].Value = p.Price;
                ws.Cells[row, 6].Value = p.StockQty;
                ws.Cells[row, 7].Value = p.Unit;
                ws.Cells[row, 8].Value = p.IsActive;
                ws.Cells[row, 9].Value = p.CreatedDate.ToString("yyyy-MM-dd HH:mm:ss");
                row++;
            }

            ws.Cells[ws.Dimension.Address].AutoFitColumns();

            using var stream = new MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "products.xlsx"
            );
        }

        public IActionResult OnGetWord(string xdata)
        {
            var stream = new MemoryStream();
            using (var wordDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();

                // Set default font and size in styles
                var stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylePart.Styles = CreateDefaultStyles();

                var body = mainPart.Document.AppendChild(new Body());

                // 1. Add logo image (top left)
                var imagePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
                if (System.IO.File.Exists(imagePath))
                {
                    var imagePart = mainPart.AddImagePart(ImagePartType.Png);
                    using (var imgStream = new FileStream(imagePath, FileMode.Open))
                    {
                        imagePart.FeedData(imgStream);
                    }
                    var element = CreateImage(mainPart.GetIdOfPart(imagePart), 120, 40); // width, height in pixels
                    var logoPara = new Paragraph(element);
                    body.AppendChild(logoPara);
                }

                // 2. Add right-aligned text box for "รหัสหน่วยงาน"
                body.AppendChild(NormalParagraph("รหัส\nหน่วยงาน", JustificationValues.Right));

                // 3. Centered, bolded titles
                body.AppendChild(CenteredBoldParagraph("สัญญาร่วมดำเนินการ", "44")); // 22pt = 44 half-points
                body.AppendChild(CenteredBoldParagraph("โครงการ.........................................", "44"));
                body.AppendChild(CenteredParagraph("ระหว่าง"));
                body.AppendChild(CenteredParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม" + xdata + ""));
                body.AppendChild(CenteredParagraph("กับ"));
                body.AppendChild(CenteredParagraph("......(ใส่ชื่อหน่วยงาน)......"));

                // 4. Main contract body (normal alignment)
                body.AppendChild(EmptyParagraph());
                body.AppendChild(NormalParagraph("สัญญาร่วมดำเนินการฉบับนี้ทำขึ้น  ณ  สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เมื่อวันที่ …………... เดือน ……….…………….. พ.ศ. …………. ระหว่าง"));
                body.AppendChild(NormalParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม  โดย......................................................ตำแหน่ง ................................................... สำนักงานตั้งอยู่เลขที่ 21 อาคารทีเอสที ทาวเวอร์ ชั้น G,17-18,23 ถนนวิภาวดีรังสิต แขวงจอมพล เขตจตุจักร กรุงเทพมหานคร 10900  ซึ่งต่อไป ในสัญญาฉบับนี้จะเรียกว่า“สสว.” ฝ่ายหนึ่ง กับ"));
                body.AppendChild(NormalParagraph("“ชื่อเต็มของหน่วยงาน” โดย     (ชื่อ - นามสกุล)       ตำแหน่ง ....................................................ผู้มีอำนาจกระทำการแทนปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ ฉบับลงวันที่...................................สำนักงานตั้งอยู่เลขที่ .....................................  ซึ่งต่อไปในสัญญาฉบับนี้จะเรียกว่า  “ชื่อหน่วยร่วม”  อีกฝ่ายหนึ่ง"));
                body.AppendChild(NormalParagraph("วัตถุประสงค์ตามสัญญาร่วมดำเนินการ"));
                body.AppendChild(NormalParagraph("คู่สัญญาทั้งสองฝ่ายมีความประสงค์ที่จะร่วมมือกันเพื่อดำเนินการภายใต้โครงการ"));
                body.AppendChild(NormalParagraph("(ชื่อโครงการที่ระบุไว้ข้างต้น)        ซึ่งต่อไปในสัญญานี้จะเรียกว่า “โครงการ”  โดยมีรายละเอียดโครงการ แผนการดำเนินงาน แผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) และบรรดาเอกสารแนบท้ายสัญญาฉบับนี้  ซึ่งให้ถือเป็นส่วนหนึ่งของสัญญาฉบับนี้  มีระยะเวลาตั้งแต่วันที่..................................จนถึงวันที่......................................โดยมีวัตถุประสงค์ในการดำเนินโครงการ  ดังนี้"));
                body.AppendChild(NormalParagraph("1. ……………………………………………………………"));
                body.AppendChild(NormalParagraph("2. …………………………………………………..………."));
                body.AppendChild(NormalParagraph("3. ……………………………………………………………"));

                // --- PAGE BREAK ---
                body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

                // --- PAGE 2: Content ---
                body.AppendChild(BoldUnderlineParagraph("ข้อ 1   ขอบเขตหน้าที่ของ “สสว.”"));
                body.AppendChild(NormalParagraph("1.1  คณะร่วมดำเนินการโครงการโดยสนับสนุนงบประมาณ จำนวน ...............บาท ..."));
                body.AppendChild(NormalParagraph("1.2  ประสานการดำเนินโครงการ เพื่อให้บรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์"));
                body.AppendChild(NormalParagraph("1.3  กำกับ ติดตามและประเมินผลการดำเนินงานของโครงการ"));

                body.AppendChild(EmptyParagraph());

                body.AppendChild(BoldUnderlineParagraph("ข้อ 2   ขอบเขตหน้าที่ของ “ชื่อหน่วยร่วม”"));
                body.AppendChild(NormalParagraph("2.1  คณะร่วมดำเนินการโครงการตามวัตถุประสงค์ของโครงการและขอบเขตการดำเนินการ ..."));
                body.AppendChild(BoldParagraph("2.2  ต้องดำเนินโครงการ"));
                body.AppendChild(NormalParagraph("ปฏิบัติตามแผนการดำเนินงาน ..."));
                body.AppendChild(BoldParagraph("2.3  ต้องประสานการดำเนินโครงการ"));
                body.AppendChild(NormalParagraph("เพื่อให้บรรลุวัตถุประสงค์ ..."));
                body.AppendChild(NormalParagraph("2.4  ต้องให้ความร่วมมือกับ สสว. ในการกำกับ ติดตาม ..."));

                body.AppendChild(EmptyParagraph());

                body.AppendChild(BoldUnderlineParagraph("ข้อ 3   อื่น ๆ"));
                body.AppendChild(NormalParagraph("3.1  หากผู้มีอำนาจลงนามฝ่ายหนึ่งประสงค์จะขอถอนตัว ..."));
                body.AppendChild(NormalParagraph("3.2  หากผู้มีอำนาจลงนามฝ่ายหนึ่งประสงค์จะขอขยายระยะเวลา ..."));


                // --- Add header with running page number ---
                var headerPart = mainPart.AddNewPart<HeaderPart>();
                string headerPartId = mainPart.GetIdOfPart(headerPart);
                headerPart.Header = new Header(
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification() { Val = JustificationValues.Right }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                            ),
                            new FieldChar() { FieldCharType = FieldCharValues.Begin }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                            ),
                            new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                            ),
                            new FieldChar() { FieldCharType = FieldCharValues.Separate }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                            ),
                            new Text("1")
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                            ),
                            new FieldChar() { FieldCharType = FieldCharValues.End }
                        )
                    )
                );

                var footerPart = mainPart.AddNewPart<FooterPart>();
                string footerPartId = mainPart.GetIdOfPart(footerPart);
                footerPart.Footer = new Footer(
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification() { Val = JustificationValues.Center }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                            ),
                            new FieldChar() { FieldCharType = FieldCharValues.Begin }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                            ),
                            new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                            ),
                            new FieldChar() { FieldCharType = FieldCharValues.Separate }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                            ),
                            new Text("1")
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                            ),
                            new FieldChar() { FieldCharType = FieldCharValues.End }
                        )
                    )
                );

                var sectionProps = new SectionProperties(
                    new HeaderReference() { Type = HeaderFooterValues.Default, Id = headerPartId },
                    new FooterReference() { Type = HeaderFooterValues.Default, Id = footerPartId },
                    new PageSize() { Width = 11906, Height = 16838 }, // A4 size
                    new PageMargin() { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720, Gutter = 0 }
                );
                body.AppendChild(sectionProps);
            }
            stream.Position = 0;
            return File(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "contract.docx");
        }

        // Helper: Create default styles for TH SarabunPSK 16pt
        private static Styles CreateDefaultStyles()
        {
            return new Styles(
                new Style(
                    new StyleName() { Val = "Normal" },
                    new BasedOn() { Val = "Normal" },
                    new UIPriority() { Val = 1 },
                    new PrimaryStyle(),
                    new StyleRunProperties(
                        new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                        new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "32" } // Correct namespace and usage // 16pt = 32 half-points
                    )
                )
            );
        }

        // Helper methods for formatting

        private static Paragraph CenteredBoldParagraph(string text, string fontSize = "32") =>
            new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                        new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = fontSize },
                        new Bold()
                    ),
                    new Text(text)
                )
            );

        private static Paragraph CenteredParagraph(string text) =>
            new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                        new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" } // Correct namespace and usage
                    ),
                    new Text(text)
                )
            );

        private static Paragraph RightParagraph(string text) =>
        new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Right }),
            new Run(
                new RunProperties(
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" } // Correct namespace and usage
                ),
                new Text(text)
            )
        );
        // Fix for CS0117: 'FontSize' does not contain a definition for 'Val'
        // The issue arises because the incorrect namespace or type is being used for FontSize.
        // Replace the problematic line with the correct usage of FontSize from DocumentFormat.OpenXml.Wordprocessing.

        private static Paragraph NormalParagraph(string text, JustificationValues? align = null, string fontSize = null) =>
            align != null
                ? new Paragraph(
                    new ParagraphProperties(new Justification { Val = align.Value }),
                    new Run(
                        new RunProperties(
                            new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                            new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = fontSize }
                        ),
                        new Text(text)
                    )
                )
                : new Paragraph(
                    new Run(
                        new RunProperties(
                            new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                            new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = fontSize }
                        ),
                        new Text(text)
                    )
                );
        private static Paragraph EmptyParagraph() =>
            new Paragraph(new Run(
                new RunProperties(
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "32" } // Correct namespace and usage
                ),
                new Text("")
            ));

        private static Paragraph BoldUnderlineParagraph(string text) =>
            new Paragraph(
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                        new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "32" }, // Correct namespace and usage,
                        new Bold(),
                        new Underline { Val = UnderlineValues.Single }
                    ),
                    new Text(text)
                )
            );

        private static Paragraph BoldParagraph(string text) =>
            new Paragraph(
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                        new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "32" }, // Correct namespace and usage,
                        new Bold()
                    ),
                    new Text(text)
                )
            );
        // Helper: Create a paragraph that starts halfway down the page

        // Helper for image insertion
        private static Drawing CreateImage(string relationshipId, long widthPx, long heightPx)
        {
            const long emusPerInch = 914400;
            const int pixelsPerInch = 96;
            long widthEmus = widthPx * emusPerInch / pixelsPerInch;
            long heightEmus = heightPx * emusPerInch / pixelsPerInch;

            return new Drawing(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = widthEmus, Cy = heightEmus },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties
                    {
                        Id = (UInt32Value)1U,
                        Name = "Picture 1"
                    },
                    new DocumentFormat.OpenXml.Drawing.Graphic(
                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties
                                    {
                                        Id = (UInt32Value)0U,
                                        Name = "New Bitmap Image.jpg"
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()
                                ),
                                new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                    new DocumentFormat.OpenXml.Drawing.Blip
                                    {
                                        Embed = relationshipId,
                                        CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Stretch(
                                        new DocumentFormat.OpenXml.Drawing.FillRectangle()
                                    )
                                ),
                                new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                    new DocumentFormat.OpenXml.Drawing.Transform2D(
                                        new DocumentFormat.OpenXml.Drawing.Offset { X = 0L, Y = 0L },
                                        new DocumentFormat.OpenXml.Drawing.Extents { Cx = widthEmus, Cy = heightEmus }
                                    ),
                                    new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                        new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                    )
                                    { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }
                                )
                            )
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
            );
        }





        #region สสว. สัญญารับเงินอุดหนุน
        // This is your specific handler for the contract report
        public IActionResult OnGetWordContactAllowance()
        {

            var wordBytes = _AllowanceService.OnGetWordContact_Allowance();
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญารับเงินอุดหนุน.docx");


        }
        // Helper for colored, bold, centered paragraph
        #endregion สสว. สัญญารับเงินอุดหนุน



        #region สสว. สัญญาเงินกู้ยืม โครงการพลิกฟื้นวิสาห 
        public IActionResult OnGetWordContactBorrowMoney()
        {

            var wordBytes = _BorrowMoneyService.OnGetWordContact_orrowMoney();
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาเงินกู้ยืมโครงการพลิกฟื้นวิสาห.docx");

        }
        #endregion สสว. สัญญาเงินกู้ยืม โครงการพลิกฟื้นวิสาห 

        #region  4.1.3.3. สัญญาจ้างลูกจ้าง EC
        public async Task<IActionResult> OnGetWordContact_EC(string ContractId="3")
        {
            var wordBytes = await _HireEmployee.OnGetWordContact_HireEmployee(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างลูกจ้าง.docx");

        }

        public async Task OnGetWordContact_EC_PDF(string ContractId = "1")
        {
            var wordBytes = await _HireEmployee.OnGetWordContact_HireEmployee_ToPDF(ContractId, "EC");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "EC");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "EC_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }
        #endregion 4.1.3.3. สัญญาจ้างลูกจ้าง

        #region 4.1.1.2.15.สัญญาจ้างทำของ CWA
        public async Task<IActionResult> OnGetWordContact_CWA(string ContractId="1")
        {
            var wordBytes = await _ContactToDoThingService.OnGetWordContact_ToDoThing(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างทำของ.docx");
        }

        public async Task OnGetWordContact_CWA_PDF(string ContractId = "1")
        {
            var wordBytes = await _ContactToDoThingService.OnGetWordContact_ToDoThing_ToPDF(ContractId, "CWA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CWA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "CWA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }
        #endregion 4.1.1.2.15.สัญญาจ้างทำของ CWA

        #region 4.1.1.2.14.สัญญาจ้างผู้เชี่ยวชาญรายบุคคลหรือจ้างบริษัทที่ปรึกษา ร.317-60 CTR31760
        public async Task<IActionResult> OnGetWordContact_CTR31760(string ContractId ="1")
        {
            var wordBytes =await _ConsultantService.OnGetWordContact_ConsultantService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างที่ปรึกษา.docx");
        }
        public async Task OnGetWordContact_CTR31760_PDF(string ContractId = "1")
        {
            var wordBytes = await _ConsultantService.OnGetWordContact_ConsultantService_ToPDF(ContractId, "CTR31760");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CTR31760");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "CTR31760_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }
        #endregion 4.1.1.2.14.สัญญาจ้างที่ปรึกษา CTR31760

        #region 4.1.1.2.13.สัญญาเช่าเครื่องถ่ายเอกสาร ร.314-60 PML31460

        public async Task<IActionResult> OnGetWordContact_PML31460(string ContractId ="1")
        {
            var wordBytes = await _wordEContract_LoanPrinterService.OnGetWordContact_LoanPrinter(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาเช่าเครื่องถ่ายเอกสาร ร.314-60.docx");
        }
        public async Task OnGetWordContact_PML31460_PDF(string ContractId = "1")
        {
            var wordBytes = await _wordEContract_LoanPrinterService.OnGetWordContact_LoanPrinter_ToPDF(ContractId, "PML31460");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PML31460");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "PML31460_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }

        #endregion 4.1.1.2.13.สัญญาเช่าเครื่องถ่ายเอกสาร ร.314-60 PML31460

        #region 4.1.1.2.12.สัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ร.310-60 SMC31060
        public async Task<IActionResult> OnGetWordContact_SMC31060(string ContractId ="1")
        {
            var wordBytes = await _maintenanceComputerService.OnGetWordContact_MaintenanceComputer(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ร.310-60.docx");
        }
        public async Task OnGetWordContact_SMC31060_PDF(string ContractId = "1")
        {
            var wordBytes = await _maintenanceComputerService.OnGetWordContact_MaintenanceComputer_ToPDF(ContractId,"SMC31060");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SMC31060");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "SMC31060_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }


        #endregion 4.1.1.2.12.สัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ร.310-60 SMC31060

        #region 4.1.1.2.11.สัญญาเช่าคอมพิวเตอร์ ร.309-60 CLA30960
        public async Task<IActionResult> OnGetWordContact_CLA30960(string ContractId ="1")
        {
            var wordBytes =await _LoanComputerService.OnGetWordContact_LoanComputer(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาเช่าคอมพิวเตอร์ ร.309-60.docx");
        }
        public async Task OnGetWordContact_CLA30960_PDF(string ContractId = "1")
        {
            var wordBytes = await _LoanComputerService.OnGetWordContact_LoanComputer_ToPDF(ContractId, "CLA30960");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CLA30960");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "CLA30960_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }
        #endregion 4.1.1.2.11.สัญญาเช่าคอมพิวเตอร์ ร.309-60 CLA30960

        #region 4.1.1.2.10.สัญญาซื้อขายและอนุญาตให้ใช้สิทธิในโปรแกรมคอมพิวเตอร์ ร.308-60 SLA30860
        public async Task<IActionResult> OnGetWordContact_SLA30860(string ContractId ="1")
        {
            var wordBytes = await _BuyAgreeProgram.OnGetWordContact_BuyAgreeProgram(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาซื้อขายและอนุญาตให้ใช้สิทธิในโปรแกรมคอมพิวเตอร์.docx");
        }
        public async Task OnGetWordContact_SLA30860_PDF(string ContractId = "1")
        {
            var wordBytes = await _BuyAgreeProgram.OnGetWordContact_BuyAgreeProgram_ToPDF(ContractId, "SLA30860");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SLA30860");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "SLA30860_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }


        #endregion 4.1.1.2.10.สัญญาซื้อขายและอนุญาตให้ใช้สิทธิในโปรแกรมคอมพิวเตอร์ ร.308-60 SLA30860

        #region 4.1.1.2.9.สัญญาซื้อขายคอมพิวเตอร์ CPA
        public async Task<IActionResult> OnGetWordContact_CPA(string ContractId ="1")
        {
            var wordBytes = await _BuyOrSellComputerService.OnGetWordContact_BuyOrSellComputerService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาซื้อขายคอมพิวเตอร์.docx");
        }
        public async Task OnGetWordContact_CPA_PDF(string ContractId = "1")
        {
            var wordBytes = await _BuyOrSellComputerService.OnGetWordContact_BuyOrSellComputerService_ToPDF(ContractId);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CPA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "CPA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            //return File(wordBytes, "application/pdf", "CPA_" + ContractId + ".pdf");
        }


        #endregion 4.1.1.2.9.สัญญาซื้อขายคอมพิวเตอร์ CPA

        #region 4.1.1.2.8.สัญญาซื้อขาย ร.305-60 SPA30560

        public async Task<IActionResult> OnGetWordContact_SPA30560(string ContractId = "1")
        {
            var wordBytes = await _BuyOrSellService.OnGetWordContact_BuyOrSellService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาซื้อขาย.docx");
        }

        public async Task OnGetWordContact_SPA30560_PDF(string ContractId = "1")
        {
            var wordBytes = await _BuyOrSellService.OnGetWordContact_BuyOrSellService_ToPDF(ContractId, "SPA30560");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SPA30560");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "SPA30560_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }

        #endregion 4.1.1.2.8.สัญญาซื้อขาย ร.305-60 SPA30560

        #region 4.1.1.2.7.สัญญาการรักษาข้อมูลที่เป็นความลับ NDA
        public async Task<IActionResult> OnGetWordContact_NDA(string ContractId ="1")
        {
            var wordBytes = await _DataSecretService.OnGetWordContact_DataSecretService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาการรักษาข้อมูลที่เป็นความลับ.docx");
        }
        public async Task OnGetWordContact_NDA_PDF(string ContractId = "1")
        {
            var wordBytes = await _DataSecretService.OnGetWordContact_DataSecretService_ToPDF(ContractId,"NDA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "NDA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "NDA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล.pdf");
        }

        #endregion 4.1.1.2.7.สัญญาการรักษาข้อมูลที่เป็นความลับ NDA

        #region 4.1.1.2.6.บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล PDSA
        public async Task<IActionResult> OnGetWordContact_PDSA(string ContractId ="3")
        {
            var wordBytes = await _DataPersonalService.OnGetWordContact_DataPersonalService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล.docx");
        }
        public async Task OnGetWordContact_PDSA_PDF(string ContractId = "3")
        {
            var wordBytes = await _DataPersonalService.OnGetWordContact_DataPersonalService_ToPDF(ContractId,"PDSA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDSA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "PDSA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

           // return File(wordBytes, "application/pdf", "บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล.pdf");
        }
        # endregion 4.1.1.2.6.บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล PDSA

        #region 4.1.1.2.5.บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วมตัวอย่างหน้าจอ JDCA

        public async Task<IActionResult> OnGetWordContact_JDCA(string ContractId="1")
        {
            var wordBytes = await _ControlDataService.OnGetWordContact_ControlDataService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วม.docx");

        }
        public async Task OnGetWordContact_JDCA_PDF(string ContractId = "1")
        {
            var wordBytes = await _ControlDataService.OnGetWordContact_ControlDataServiceHtmlToPdf(ContractId,"JDCA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JDCA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "JDCA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);
            // return File(wordBytes, "application/pdf", "บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วม.pdf");
        }
        #endregion 4.1.1.2.5.บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วมตัวอย่างหน้าจอ JDCA


        #region 4.1.1.2.4.บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล PDPA
        public async Task<IActionResult> OnGetWordContact_PDPA(string ContractId = "1")
        {
            var wordBytes = await _PersernalProcessService.OnGetWordContact_PersernalProcessService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล.docx");
        }
        public async Task OnGetWordContact_PDPA_PDF(string ContractId = "1")
        {
            var wordBytes = await _PersernalProcessService.OnGetWordContact_PersernalProcessService_HtmlToPDF(ContractId, "PDPA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDPA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "PDPA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);
            // return File(wordBytes, "application/pdf", "บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล.pdf");
        }
        #endregion 4.1.1.2.4.บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล PDPA

        #region 4.1.1.2.3.บันทึกข้อตกลงความร่วมมือ MOU
        public async Task<IActionResult> OnGetWordContact_MOU(string ContractId = "2")
        {
            var wordBytes = await _MemorandumService.OnGetWordContact_MemorandumService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "บันทึกข้อตกลงความร่วมมือ.docx");
        }
        public async Task OnGetWordContact_MOU_PDF(string ContractId = "3")
        {
            var wordBytes = await _MemorandumService.OnGetWordContact_MemorandumService_HtmlToPDF(ContractId,"MOU");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOU");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "MOU_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);
            //   return File(wordBytes, "application/pdf", "MOU_" + ContractId + ".pdf");
        }
        #endregion  4.1.1.2.3.บันทึกข้อตกลงความร่วมมือ MOU

        #region 4.1.1.2.2.สัญญารับเงินอุดหนุน GA
        public async Task<IActionResult> OnGetWordContact_GA(string ContractId = "1")
        {
            var wordBytes = await _SupportSMEsService.OnGetWordContact_SupportSMEsService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "บันทึกข้อตกลงความร่วมมือในการสนับสนุน SMEs.docx");
        }
        public async Task OnGetWordContact_GA_PDF(string ContractId = "1")
        {
            var pdfBytes = await _SupportSMEsService.OnGetWordContact_SupportSMEsService_HtmlToPDF(ContractId,"GA");

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "GA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "GA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);
           // return File(pdfBytes, "application/pdf", "GA_"+ContractId+".pdf");
        }


        #endregion 4.1.1.2.2.สัญญารับเงินอุดหนุน GA

        #region 4.1.1.2.1.สัญญาร่วมดำเนินการ JOA
        public async Task<IActionResult> OnGetWordContact_JOA(string ContractId="32")
        {
            var wordBytes = await _JointOperationService.OnGetWordContact_JointOperationService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาร่วมดำเนินการ.docx");
        }

        public async Task OnGetWordContact_JOA_PDF(string ContractId = "32")
        {
            var wordBytes = await _JointOperationService.OnGetWordContact_JointOperationServiceHtmlToPDF(ContractId);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JOA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "JOA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

          //  return File(wordBytes, "application/pdf", "JOA_" + ContractId + ".pdf");
        }
        #endregion 4.1.1.2.1.สัญญาร่วมดำเนินการ JOA



        #region Test Header logo
        public async Task<IActionResult> OnGetWordContact_TestLogo(string PDPAid = "1")
        {
            var wordBytes = await _Test_HeaderLOGOService.OnGetWordContact_PersernalProcessService(PDPAid);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล.docx");
        }

        #endregion Test Header logo


        #region Word to PDF using Interop
        public async Task<IActionResult> OnGetWordtoPDF(string ContractId = "1")
        {
            var pdfBytes = await _SupportSMEsService.OnGetWordContact_SupportSMEsService_HtmlToPDF(ContractId,"68");
            return File(pdfBytes, "application/pdf", "บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล.pdf");
        }

        #endregion Word to PDF using Interop
    }
}
