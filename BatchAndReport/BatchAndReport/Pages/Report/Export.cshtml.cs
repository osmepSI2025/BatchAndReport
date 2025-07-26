using BatchAndReport.DAO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
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
        public ExportModel(SmeDAO smeDao, WordEContract_AllowanceService allowanceService
            , WordEContract_LoanPrinterService wordEContract_LoanPrinterService
            , WordEContract_ContactToDoThingService ContactToDoThingService
            , WordEContract_HireEmployee HireEmployee
            , WordEContract_BorrowMoneyService BorrowMoneyService
            , WordEContract_MaintenanceComputerService maintenanceComputerService
            , WordEContract_LoanComputerService LoanComputerService
            , WordEContract_BuyAgreeProgram BuyAgreeProgram
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

        #region  4.1.3.3. สัญญาจ้างลูกจ้าง
        public IActionResult OnGetWordContactHireEmployee()
        {
            var wordBytes = _HireEmployee.OnGetWordContact_HireEmployee();
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างลูกจ้าง.docx");

        }
        #endregion 4.1.3.3. สัญญาจ้างลูกจ้าง

        #region 4.1.1.2.15.สัญญาจ้างทำของ
        public IActionResult OnGetWordContactToDoThing()
        {
            var wordBytes = _ContactToDoThingService.OnGetWordContact_ToDoThing();
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างทำของ.docx");


        }
        #endregion 4.1.1.2.15.สัญญาจ้างทำของ
        #region 4.1.1.2.13.สัญญาเช่าเครื่องถ่ายเอกสาร ร.314-60

        public IActionResult OnGetWordContact_LoanPrinter()
        {
            var wordBytes = _wordEContract_LoanPrinterService.OnGetWordContact_LoanPrinter();
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาเช่าเครื่องถ่ายเอกสาร ร.314-60.docx");
        }

        #endregion 4.1.1.2.13.สัญญาเช่าเครื่องถ่ายเอกสาร ร.314-60

        #region 4.1.1.2.12.สัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ร.310-60
        public IActionResult OnGetWordContact_MaintenanceComputer()
        {
            var wordBytes = _maintenanceComputerService.OnGetWordContact_MaintenanceComputer();
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ร.310-60.docx");
        }
        #endregion

        #region 4.1.1.2.11.สัญญาเช่าคอมพิวเตอร์ ร.309-60
        public IActionResult OnGetWordContact_LoanComputer()
        {
            var wordBytes = _LoanComputerService.OnGetWordContact_LoanComputer();
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาเช่าคอมพิวเตอร์ ร.309-60.docx");
        }
        #endregion 4.1.1.2.11.สัญญาเช่าคอมพิวเตอร์ ร.309-60

        #region 4.1.1.2.10.สัญญาซื้อขายและอนุญาตให้ใช้สิทธิในโปรแกรมคอมพิวเตอร์ ร.308-60
        public IActionResult OnGetWordContact_BuyAgreeProgram()
        {
            var wordBytes = _BuyAgreeProgram.OnGetWordContact_BuyAgreeProgram();
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาซื้อขายและอนุญาตให้ใช้สิทธิในโปรแกรมคอมพิวเตอร์.docx");
        }
        #endregion 4.1.1.2.10.สัญญาซื้อขายและอนุญาตให้ใช้สิทธิในโปรแกรมคอมพิวเตอร์ ร.308-60
    }
}
