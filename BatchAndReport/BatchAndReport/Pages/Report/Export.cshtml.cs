using BatchAndReport.DAO;
using BatchAndReport.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using QuestPDF.Fluent;
using System.IO;
using System.Threading.Tasks;
namespace BatchAndReport.Pages.Report
{
    public class ExportModel : PageModel
    {
        private readonly SmeDAO _smeDao;
        private readonly IWordEContract_AllowanceService _wordEContract_AllowanceService;
        public ExportModel(SmeDAO smeDao, IWordEContract_AllowanceService wordEContract_AllowanceService    )
        {
            _smeDao = smeDao;
            this._wordEContract_AllowanceService = wordEContract_AllowanceService;
        }
        public IActionResult OnGetPdf()
        {
            var wordDAO = new WordToPDFDAO(); // Create an instance of WordDAO
          var Resultpdf  =  wordDAO.OnGetPdfWithInterop(); // Call the method on the instance
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
                body.AppendChild(CenteredParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม"+ xdata + ""));
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
        private static Paragraph CenteredBoldParagraph(string text) =>
            new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                        new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "32" }, // Correct namespace and usage,
                        new Bold()
                    ),
                    new Text(text)
                )
            );

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

        private static Paragraph NormalParagraph(string text, JustificationValues? align = null) =>
            align != null
                ? new Paragraph(
                    new ParagraphProperties(new Justification { Val = align.Value }),
                    new Run(
                        new RunProperties(
                            new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                            new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                        ),
                        new Text(text)
                    )
                )
                : new Paragraph(
                    new Run(
                        new RunProperties(
                            new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                            new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
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
        private static Paragraph HalfPageParagraph(string text)
        {
            // A4 page height = 16838 twips, half = ~8419 twips
            // Set SpacingBefore to 8419 to push the paragraph halfway down
            return new Paragraph(
                new ParagraphProperties(
                    new SpacingBetweenLines { Before = "8419" }, // twips
                    new Justification { Val = JustificationValues.Center }
                ),
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                        new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "32" }
                    ),
                    new Text(text)
                )
            );
        }

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
                                    ) { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }
                                )
                            )
                        ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
            );
        }

        private static Paragraph JustifiedParagraph(string text) =>
    new Paragraph(
        new ParagraphProperties(new Justification { Val = JustificationValues.Both }),
        new Run(
            new RunProperties(
                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
            ),
            new Text(text)
        )
    );
        // Helper: Paragraph with 2 tab spaces at the start of the first line
        private static Paragraph NormalParagraphWithTabs(string text, JustificationValues? align = null)
        {
            var paragraph = new Paragraph();

            // Paragraph properties (alignment and tab stops)
            var props = new ParagraphProperties();
            if (align != null)
                props.Append(new Justification { Val = align.Value });

            // Add two tab stops (every 720 = 0.5 inch, adjust as needed)
            var tabs = new Tabs(
                new TabStop { Val = TabStopValues.Left, Position = 720 },
                new TabStop { Val = TabStopValues.Left, Position = 1440 }
            );
            props.Append(tabs);
            paragraph.Append(props);

            // Add two tab characters at the start
            var run = new Run(
                new RunProperties(
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }
                ),
                new TabChar(),
                new TabChar(),
                new Text(text)
            );
            paragraph.Append(run);

            return paragraph;
        }
        private static Paragraph NormalParagraphWithTabsColor(string text, JustificationValues? align = null, string hexColor =null)
        {
            var paragraph = new Paragraph();

            // Paragraph properties (alignment and tab stops)
            var props = new ParagraphProperties();
            if (align != null)
                props.Append(new Justification { Val = align.Value });

            // Add two tab stops (every 720 = 0.5 inch, adjust as needed)
            var tabs = new Tabs(
                new TabStop { Val = TabStopValues.Left, Position = 720 },
                new TabStop { Val = TabStopValues.Left, Position = 1440 }
            );
            props.Append(tabs);
            paragraph.Append(props);

            // Add two tab characters at the start
            var run = new Run(
                new RunProperties(
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" },
                        new Color { Val = hexColor }
                ),
                new TabChar(),
                new TabChar(),
                new Text(text)
            );
            paragraph.Append(run);

            return paragraph;
        }
        #region 
        // This is your specific handler for the contract report
        public IActionResult OnGetWordContactAllowance()
        {

            var stream = new MemoryStream();

            using (var wordDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();

                // Styles
                var stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylePart.Styles = CreateDefaultStyles();

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
                    var element = CreateImage(mainPart.GetIdOfPart(imagePart), 160, 40);
                    var logoPara = new Paragraph(
                        new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                        element
                    );
                    body.AppendChild(logoPara);
                }

                // 2. Document title and subtitle
                body.AppendChild(EmptyParagraph());
                body.AppendChild(RightParagraph("เลขที่สัญญา ............................"));
                body.AppendChild(EmptyParagraph());
                body.AppendChild(CenteredBoldColoredParagraph("สัญญารับเงินอุดหนุน", "FF0000")); // Blue
                body.AppendChild(CenteredBoldColoredParagraph("ตามแนวทางการดำเนินโครงการวิสาหกิจขนาดกลางและขนาดย่อมต่อเนื่อง", "FF0000")); // Red
                body.AppendChild(HalfPageParagraph("ที่ศูนย์ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม"));
                body.AppendChild(HalfPageParagraph("วันที่...................................................."));

                // 3. Fillable lines (using underlines)
              //  body.AppendChild(EmptyParagraph());
                body.AppendChild(NormalParagraphWithTabs("ข้าพเจ้า ...................................................................................................................."));
                body.AppendChild(JustifiedParagraph("อายุ ......... ปี สัญชาติ .................. สำนักงาน/บ้านตั้งอยู่เลขที่.................. อาคาร..........................................."));
                body.AppendChild(JustifiedParagraph("หมู่ที่...........ตรอก/ซอย..........................ถนน...........................ตำบล/แขวง.................. ..................."));
                body.AppendChild(JustifiedParagraph("เขต/อำเภอ................... จังหวัด.................ทะเบียนนิติบุคคลเลขที่/เลขประจำตัวประชาชนที่............................................"));
                body.AppendChild(JustifiedParagraph("จดทะเบียนเป็นนิติบุคคลเมื่อวันที่ .........................................."));

                // 4. Main body (sample)
                //   body.AppendChild(EmptyParagraph());
                body.AppendChild(NormalParagraph("ซึ่งต่อไปนี้จะเรียกบุคคลผู้มีนามตามที่ระบุข้างต้นทั้งหมดว่า \"ผู้รับการอุดหนุน\" ได้ทำสัญญาฉบับนี้ให้ไว้แก่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งต่อไปนี้จะเรียกว่า \"ผู้ให้การอุดหนุน\" โดยมีสาระสำคัญดังนี้"));
                body.AppendChild(NormalParagraphWithTabs("ข้อ1. ผู้รับการอุดหนุนได้ขอรับความช่วยเหลือผ่านการอุดหนุนตามมาตรการฟื้นฟูกิจการวิสาหกิจ ขนาดกลางและขนาดย่อมจากผู้ให้การอุดหนุนเป็นจำนวนเงิน ..................... บาท (...........................) ปลอดการชำระเงินต้น ................. เดือน โดยไม่มีดอกเบี้ย แต่มีภาระต้องชำระคืนเงินต้น  "));
                body.AppendChild(NormalParagraphWithTabs("ข้อ2. ผู้ให้การอุดหนุนจะให้ความช่วยเหลือด้วยการให้เงินอุดหนุนแก่ผู้รับการอุดหนุน ด้วยการนำเงินหรือโอนเงินเข้าบัญชีธนาคารกรุงไทย จำกัด (มหาชน) สาขา ..................................... เลขที่บัญชี ................................................. ชื่อบัญชี .......................... ซึ่งเป็นบัญชีของผู้รับการอุดหนุน จำนวนเงิน ........................... บาท (..............................) และให้ถือว่าผู้รับการอุดหนุนได้รับเงินอุดหนุนตามสัญญานี้ไปจากผู้ให้การอุดหนุนแล้ว ในวันที่เงินเข้าบัญชีของผู้รับการอุดหนุนดังกล่าว"));
                body.AppendChild(NormalParagraphWithTabs("ข้อ3. ห้ามผู้รับการอุดหนุนนำเงินอุดหนุนไปชำระหนี้เดิมที่มีอยู่ก่อนทำสัญญานี้"));
                body.AppendChild(NormalParagraphWithTabs("ข้อ4. ผู้รับการอุดหนุนยินยอมให้ธนาคารกรุงไทย จำกัด (มหาชน) ซึ่งกระทำการแทนผู้ให้การอุดหนุน หักเงินอุดหนุนที่จะได้จากผู้ให้การอุดหนุนเป็นค่าใช้จ่ายหรือค่าธรรมเนียมในการโอนเงินเข้าบัญชีของผู้รับการอุดหนุน ซึ่งธนาคารกรุงไทย จำกัด (มหาชน) เรียกเก็บตามระเบียบของธนาคารได้ โดยไม่ต้องบอกกล่าวหรือแจ้งให้ผู้รับการอุดหนุนทราบล่วงหน้า และให้ถือว่าผู้รับการอุดหนุนได้รับเงินตามจำนวนที่เบิกไปครบถ้วนแล้ว "));
                body.AppendChild(NormalParagraphWithTabs("ข้อ5. ผู้รับการอุดหนุนตกลงผ่อนชำระเงินต้นคืนให้แก่ผู้ให้การอุดหนุนเป็นรายเดือน (งวด) ๆ ละ ไม่น้อยกว่า ....................... บาท (.....................................) ด้วยการโอนเข้าบัญชีตามที่ระบุไว้ในข้อ 2 โดยชำระเงินต้นงวดแรกในเดือนที่ ....................... นับถัดจากวันที่ได้รับเงินอุดหนุน และงวดถัดไปทุกวันที่ .................. ของเดือนจนกว่าจะชำระเสร็จสิ้น  แต่ทั้งนี้จะต้องชำระให้เสร็จสิ้นไม่เกินกว่า .............. ปี (...........) นับแต่วันที่ได้รับเงินอุดหนุน "));
                body.AppendChild(NormalParagraphWithTabs("ข้อ6. การชำระเงินคืนตาม"));
                body.AppendChild(NormalParagraphWithTabs("ข้อ5. ผู้รับการอุดหนุนตกลงจะนำเงินเข้าบัญชีเงินฝากของผู้รับการอุดหนุน ที่เปิดบัญชีไว้กับธนาคารกรุงไทย จำกัด (มหาชน) ตามข้อ 2 โดยผู้รับการอุดหนุนยินยอมให้ ธนาคารกรุงไทย จำกัด (มหาชน) ซึ่งดำเนินการแทนผู้ให้การอุดหนุน หักเงินจากบัญชีของผู้รับการอุดหนุนดังกล่าวเพื่อชำระคืนเงินอุดหนุนแก่ผู้ให้การอุดหนุน "));
                body.AppendChild(NormalParagraphWithTabs("ข้อ6. การชำระเงินคืนตามข้อ 5 ผู้รับการอุดหนุนตกลงจะนำเงินเข้าบัญชีเงินฝากของผู้รับการอุดหนุนที่เปิดบัญชีไว้กับธนาคารกรุงไทย จำกัด (มหาชน) ตามข้อ 2 โดยผู้รับการอุดหนุนยินยอมให้ ธนาคารกรุงไทย จำกัด (มหาชน) ซึ่งดำเนินการแทนผู้ให้การอุดหนุน หักเงินจากบัญชีของผู้รับการอุดหนุนดังกล่าวเพื่อชำระคืนเงินอุดหนุนแก่ผู้ให้การอุดหนุน ในแต่ละงวดเดือน พร้อมทำการโอนเงินที่พักจากบัญชีของผู้รับการอุดหนุนนำเข้าบัญชีของผู้ให้การอุดหนุนที่เปิดบัญชีไว้กับธนาคารกรุงไทย จำกัด (มหาชน) สาขา .....องค์การตลาดเพื่อเกษตรกร (จตุจักร)..... บัญชีออมทรัพย์ ชื่อบัญชีสำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เลขที่บัญชี ....035-1-52709-5.....เพื่อชำระหนี้คืนเงินอุดหนุนแก่ผู้ให้การอุดหนุนตามข้อตกลงในแต่ละงวดเดือน "));
                body.AppendChild(NormalParagraphWithTabs("ไม่ว่าผู้รับการอุดหนุนจะได้จัดทำหนังสือยินยอมให้ธนาคารกรุงไทย จำกัด (มหาชน) หักบัญชีเงินฝาก ตามวรรคหนึ่งหรือไม่ก็ตาม โดยสัญญนี้ผู้รับการอุดหนุนให้ถือว่าเป็นการทำหนังสือยินยอมให้ธนาคารกรุงไทย จำกัด (มหาชน) หักบัญชีเงินฝากตามวรรคหนึ่งด้วย "));
                body.AppendChild(NormalParagraphWithTabs("ข้อ7. ผู้รับการอุดหนุนตกลงยินยอมให้ธนาคารกรุงไทย จจำกัด (มหาชน) ซึ่งกระทำการแทนผู้ให้การอุดหนุนหักเงินที่ผู้รับการอุดหนุนได้โอนเข้าบัญชีตามข้อ 2 เพื่อชำระคืนเป็นค่าใช้จ่ายหรือค่าธรรมเนียมในการโอนเงิน"));
                body.AppendChild(NormalParagraphWithTabs("ข้อ8. ในระหว่างและตลอดระยะเวลาการตามสัญญาฉบับนี้ ผู้รับการอุดหนุนจะต้องรายงานผลการประกอบกิจการมายังผู้ให้การอุดหนุนหรือศูนย์ให้บริหร SMEs ครบวงจร ในจังหวัดที่ผู้รับการอุดหนุนมีภูมิลำเนาอยู่หรือพื้นที่ใกล้เคียงหรือหน่วยงานอื่นใดที่ผู้ให้การอุดหนุนมอบหมาย ตามหลักเกณฑ์และวิธีการที่ผู้ให้การอุดหนุนกำหนด ไม่น้อยกว่าเดือนละหนึ่งครั้ง"));
                body.AppendChild(NormalParagraphWithTabs("ข้อ9. กรณีต่อไปนี้ให้ถือว่าผู้รับการอุดหนุนปฏิบัติผิดสัญญา"));
                body.AppendChild(NormalParagraphWithTabs("9.1 ผู้รับการอุดหนุนผิดนัดชำระคืนเงินอุดหนุนไม่ว่างวดหนึ่งวดใดก็ตาม หรือไม่ชำระคืนเงินอุดหนุนภายในกำหนดระยะเวลาที่กำหนดในสัญญานี้ หรือเงินจำนวนอื่นใดที่ต้องชำระตามสัญญาฉบับนี้"));
                body.AppendChild(NormalParagraphWithTabs("9.2 ผู้รับการอุดหนุนใช้เงินอุดหนุนผิดไปจากเงื่อนไขตามสัญญา หรือผิดสัญญาแม้ข้อใดข้อหนึ่ง หรือไม่รายงานการดำเนินธุรกิจให้ผู้ให้การอุดหนุนทราบตามข้อ 8 หรือตรวจสอบในภายหลังแล้วพบว่ามีการแจ้งคุณสมบัติ หรือส่งเอกสารเป็นเท็จแก่ผู้ให้การอุดหนุน ๆ มีสิทธิบอกเลิกสัญญาได้ "));
                body.AppendChild(NormalParagraphWithTabs("ข้อ10. อื่นๆ "));
                body.AppendChild(NormalParagraphWithTabs("10.1 ในระหว่างและตลอดระยะเวลาตามสัญญานี้ ผู้รับการอุดหนุนยินยอมให้ผู้ให้การอุดหนุน หรือตัวแทนผู้ให้การอุดหนุนเข้าไปตรวจสอบติดตามการดำเนินธุรกิจ ตลอดจนเอกสารหลักฐานทางบัญชีของกิจการ สรรพเอกสารอื่น ๆ ของผู้รับการอุดหนุนได้ตลอด"));
                body.AppendChild(NormalParagraphWithTabs("10.2 คู่สัญญาตกลงให้ถือเอาเอกสารที่แนบท้ายสัญญานี้ บันทึกข้อตกลง และบรรดาข้อสัญญาต่าง ๆ  เป็นส่วนหนึ่งของสัญญานี้ที่มีผลผูกพันให้ผู้รับการอุดหนุนจะต้องปฏิบัติตาม ซึ่งเอกสารแนบท้ายนี้อาจจะทำเพิ่มเติมในภายหลังจากวันทำสัญญานี้ โดยให้ถือเป็นส่วนหนึ่งของสัญญานี้เช่นกัน และหากเอกสารแนบท้ายสัญญาขัดหรือแย้งกันผู้รับการอุดหนุนตกลงปฏิบัติตามคำวินิจฉัยของผู้ให้การอุดหนุน "));
                body.AppendChild(NormalParagraphWithTabs("10.3 บรรดาหนังสือ จดหมาย คำบอกกล่าวใด ๆ เช่น การทวงถาม การบอกเลิกสัญญา ของผู้ให้การอุดหนุนหรือผู้ที่ได้รับมอบหมายส่งไปยังสถานที่ที่ระบุไว้เป็นที่อยู่ของผู้รับการอุดหนุนข้างต้น หรือสถานที่อยู่ที่ผู้รับการอุดหนุนแจ้งเปลี่ยนแปลง โดยส่งเองหรือส่งทางไปรษณีย์ลงทะเบียน หรือไม่ลงทะเบียน ไม่ว่าจะมีผู้รับไว้ หรือไม่มีผู้ใดยอมรับไว้ หรือส่งไม่ได้เพราะผู้รับการอุดหนุนย้ายสถานที่อยู่ไปโดยมิได้แจ้งให้ผู้ให้การอุดหนุนทราบหรือหาไม่พบ หรือถูกรื้อถอนทำลายทุก ๆ กรณีดังกล่าวให้ถือว่าผู้รับการอุดหนุนได้รับโดยชอบแล้ว"));
                body.AppendChild(NormalParagraphWithTabs("10.4 การสละสิทธิ์ตามสัญญานี้ ในคราวหนึ่งคราวใดของผู้ให้การอุดหนุน หรือการที่ผู้ให้การอุดหนุนมิได้ ใช้สิทธิ์ที่มีอยู่ ไม่ถือเป็นการสละสิทธิ์ของผู้ให้การอุดหนุนในคราวต่อไปและไม่มีผลกระทบต่อการใช้สิทธิของผู้ให้การอุดหนุน ในคราวต่อไป "));
                body.AppendChild(NormalParagraphWithTabs("10.5 หากข้อกำหนด และ/หรือเงื่อนไขข้อใดข้อหนึ่งของสัญญานี้ตกเป็นโมฆะ หรือใช้บังคับไม่ได้ตามกฎหมาย ให้ข้อกำหนดและเงื่อนไขอื่น ๆ ยังคงมีผลใช้บังคับได้ต่อไปได้ โดยแยกต่างหากจากส่วนที่เป็นโมฆะหรือไม่สมบูรณ์นั้น"));
                body.AppendChild(NormalParagraphWithTabs("สัญญานี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน คู่สัญญาทั้งสองฝ่ายได้ตรวจ อ่าน และเข้าใจข้อความในสัญญานี้โดยละเอียดแล้ว เห็นว่าถูกต้องตามเจตนาทุกประการ จึงได้ลงลายมือชื่อพร้อมประทับตรา (ถ้ามี) ไว้เป็นสำคัญ ต่อหน้าพยาน ณ วัน เดือน ปี ที่ระบุไว้ข้างต้น "));

                body.AppendChild(CenteredParagraph("ลงชื่อ........................................................................ผู้ให้การอุดหนุน"));
                body.AppendChild(CenteredParagraph("(................................................................................)"));
                body.AppendChild(CenteredParagraph("ลงชื่อ.......................................................................ผู้รับการอุดหนุน"));
                body.AppendChild(CenteredParagraph("(................................................................................)"));
                body.AppendChild(CenteredParagraph("ลงชื่อ.......................................................................คู่สมรสให้ความยินยอม"));
                body.AppendChild(CenteredParagraph("(...............................................................................)"));
                body.AppendChild(CenteredParagraph("ลงชื่อ......................................................................พยาน"));
                body.AppendChild(CenteredParagraph("(...............................................................................)"));
                body.AppendChild(CenteredParagraph("ลงชื่อ.....................................................................พยาน"));
                body.AppendChild(CenteredParagraph("(...............................................................................)"));

                // next page
                body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

                body.AppendChild(CenteredParagraph("คำรับรองสถานภาพการสมรส"));
                body.AppendChild(NormalParagraphWithTabs("ข้าพเจ้า………………………………………………………………………………………………………………………………………….\r\nขอรับรองว่าสถานภาพการสมรสของข้าพเจ้าปัจจุบันมีสถานะ\r\n"));
                body.AppendChild(NormalParagraphWithTabs("“ข้าพเจ้าขอรับรองว่าสถานภาพการสมรสที่แจ้งในหนังสือฉบับนี้เป็นความจริงทุกประการหากไม่เป็นความจริงแล้ว ความเสียหายใด ๆ ที่จะเกิดกับผู้ให้การอุดหนุน ข้าพเจ้ายินยอมรับผิดชดใช้ให้แก่ผู้ให้การอุดหนุนทั้งสิ้น”"));
                body.AppendChild(CenteredParagraph("ลงชื่อ.............................................................รับรอง"));
                body.AppendChild(CenteredParagraph("(............................................................)"));
                body.AppendChild(CenteredParagraph("ลงชื่อ....................................................พยาน                    ลงชื่อ ........................................................พยาน"));
                body.AppendChild(CenteredParagraph("(............................................................)                                 (.........................................................)"));

                body.AppendChild(EmptyParagraph());
                body.AppendChild(RightParagraph("........................................................./ผู้พิมพ์"));
                body.AppendChild(RightParagraph("........................................................./ผู้ตรวจ"));


                // --- Add header for first page (empty) ---
                var firstHeaderPart = mainPart.AddNewPart<HeaderPart>();
                string firstHeaderPartId = mainPart.GetIdOfPart(firstHeaderPart);
                firstHeaderPart.Header = new Header(
                    new Paragraph() // Empty paragraph, so no page number on first page
                );

                // --- Add header for other pages (centered page number) ---
                var headerPart = mainPart.AddNewPart<HeaderPart>();
                string headerPartId = mainPart.GetIdOfPart(headerPart);
                headerPart.Header = new Header(
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
                    new HeaderReference() { Type = HeaderFooterValues.First, Id = firstHeaderPartId },
                    new HeaderReference() { Type = HeaderFooterValues.Default, Id = headerPartId },
                    new PageSize() { Width = 11906, Height = 16838 }, // A4 size
                    new PageMargin() { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720, Gutter = 0 },
                    new TitlePage() // This enables different first page header/footer
                );
                body.AppendChild(sectionProps);
            }
            stream.Position = 0;
            return File(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สสว.สัญญารับเงินอุดหนุน.docx");
        }
        // Helper for colored, bold, centered paragraph


        #region สัญญากู้ยืมเงิน
        public IActionResult OnGetWordContactBorrowMoney()
        {

            var stream = new MemoryStream();

            using (var wordDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();

                // Styles
                var stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylePart.Styles = CreateDefaultStyles();

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
                    var element = CreateImage(mainPart.GetIdOfPart(imagePart), 160, 40);
                    var logoPara = new Paragraph(
                        new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                        element
                    );
                    body.AppendChild(logoPara);
                }

              // 2. Document title and subtitle
        body.AppendChild(EmptyParagraph());
        body.AppendChild(RightParagraph("ทะเบียนลูกค้า ............................"));
        body.AppendChild(RightParagraph("เลขที่สัญญา ............................"));
        body.AppendChild(EmptyParagraph());
        body.AppendChild(CenteredBoldColoredParagraph("สัญญากู้ยืมเงิน","FF0000")); // Blue
        body.AppendChild(CenteredBoldColoredParagraph("โครงการเงินทุนพลิกฟื้นวิสาหกิจขนาดย่อม","FF0000")); // Red
        body.AppendChild(RightParagraph("ทำที่ ธนาคารพัฒนาวิสาหกิจขนาดกลางและขนาดย่อมแห่งประเทศไทย"));
        body.AppendChild(RightParagraph("สำนักงานใหญ่/สาขา.........................................................."));

        // 3. Fillable lines (using underlines)
        // body.AppendChild(EmptyParagraph());
        body.AppendChild(NormalParagraphWithTabs("ข้าพเจ้า ...................................................................................................................."));
        body.AppendChild(JustifiedParagraph("อายุ ......... ปี สัญชาติ .................. สำนักงาน/บ้านตั้งอยู่เลขที่.................. อาคาร..........................................."));
        body.AppendChild(JustifiedParagraph("หมู่ที่...........ตรอก/ซอย..........................ถนน...........................ตำบล/แขวง.................. ..................."));
        body.AppendChild(JustifiedParagraph("เขต/อำเภอ................... จังหวัด.................ทะเบียนนิติบุคคลเลขที่/เลขประจำตัวประชาชนที่............................................"));
        body.AppendChild(JustifiedParagraph("จดทะเบียนเป็นนิติบุคคลเมื่อวันที่..................................ซึ่งต่อไปนี้จะเรียกบุคคลผู้มีนามตามที่ระบุข้างต้นทั้งหมดว่า \"ผู้กู้\"ได้ทำสัญญาฉบับนี้ให้ไว้แก่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม" +
         "ซึ่งต่อไปนี้จะเรียกว่า \"ผู้ให้กู้\" โดยมีสาระสำคัญดังนี้"));

        // 4. Main body (sample)
        //  body.AppendChild(EmptyParagraph());
        body.AppendChild(NormalParagraph("ซึ่งต่อไปนี้จะเรียกบุคคลผู้มีนามตามที่ระบุข้างต้นทั้งหมดว่า \"ผู้รับการอุดหนุน\" ได้ทำสัญญาฉบับนี้ให้ไว้แก่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งต่อไปนี้จะเรียกว่า \"ผู้ให้การอุดหนุน\" โดยมีสาระสำคัญดังนี้"));
        body.AppendChild(NormalParagraphWithTabs("ข้อ 1. วัตถุประสงค์และวงเงินกู้"));
        body.AppendChild(NormalParagraphWithTabs("โดยผู้กู้ได้กู้เงินจากผู้ให้กู้เป็นจำนวนเงิน.....................................บาท(....................................................)เ" +
         "พื่อนำไปใช้จ่ายเป็นเงินทุนหมุนเวียน"));
        body.AppendChild(NormalParagraphWithTabsColor("โดยไม่นำเงินที่กู้ยืมไปชำระหนี้ที่มีอยู่ก่อนยื่นคำขอกู้ยืมเงิน",null,"FFF0000"));

       body.AppendChild(NormalParagraphWithTabs("กำหนดชำระเงินกู้เสร็จสิ้นภายใน...........ปี...........เดือน  โดยมีระยะเวลาปลอดเงินต้น................เดือน"));
        body.AppendChild(NormalParagraphWithTabs("ข้อ 2. การเบิกจ่ายเงินกู้"));
        body.AppendChild(NormalParagraphWithTabs("ผู้ให้กู้จะจ่ายเงินกู้แก่ผู้กู้ตามเงื่อนไขการใช้เงินกู้ในข้อ 1.และตามรายละเอียดการใช้เงินกู้ ซึ่งผู้กู้ได้แจ้งไว้ในคำขอสินเชื่อและเอกสารแนบท้ายคำขอสินเชื่อโดยถือเป็นส่วนหนึ่งของสัญญากู้เงินฉบับนี้ด้วย" +
         "หากปรากฏว่ารายการขอเบิกเงินกู้งวดใดไม่เป็นไปตามเงื่อนไขและรายละเอียดดังกล่าว " +
         "เป็นสิทธิของผู้ให้กู้แต่ฝ่ายเดียวที่จะพิจารณาไม่ให้เบิกเงินกู้ก็ได้"));
        body.AppendChild(NormalParagraphWithTabsColor("โดยผู้ให้กู้จะจ่ายเงินกู้ให้ผู้กู้ด้วยการนำเงินหรือโอนเงินเข้าบัญชีที่ ธนาคารกรุงไทย จำกัด (มหาชน)" +
         "\r\nสาขา..............................................ชื่อบัญชี................................................................................ซึ่งเป็นบัญชีของผู้กู้" +
         "\r\nเลขที่บัญชี...................................จำนวนเงิน...........................บาท (.............................................)" +
         "และให้ถือว่าผู้กู้ได้รับเงินกู้ตามสัญญานี้ไปจากผู้ให้กู้แล้วในวันที่เงินเข้าบัญชีของผู้กู้ดังกล่าว\r\n"));
        body.AppendChild(NormalParagraphWithTabsColor("ทั้งนี้ ผู้กู้ยินยอมให้ผู้ให้กู้ หรือ ธนาคารพัฒนาวิสาหกิจขนาดกลางและขนาดย่อมแห่งประเทศไทย" +
         "ซึ่งกระทำการแทนผู้ให้กู้ หักเงินจากจำนวนเงินกู้ที่ผู้กู้ขอเบิกจากผู้ให้กู้เป็นค่าวิเคราะห์โครงการ ค่าอากรแสตมป์ ค่าใช้จ่ายหรือค่าธรรมเนียมในการโอนเงินเข้าบัญชีของผู้กู้ซึ่งธนาคารกรุงไทย จำกัด (มหาชน)" +
         " เรียกเก็บตามระเบียบของธนาคาร โดยไม่ต้องบอกกล่าวหรือแจ้งให้ผู้กู้ทราบ" +
         "โดยให้ถือว่าผู้กู้ได้รับเงินกู้ตามจำนวนที่เบิกไปครบถ้วนแล้วและสละสิทธิ์ที่จะเรียกร้องอย่างใด ๆ" +
         "ต่อผู้ให้กู้และหรือธนาคารพัฒนาวิสาหกิจขนาดกลางและขนาดย่อมแห่งประเทศไทย ที่ดำเนินการแทนตามที่ได้รับมอบหมายจากผู้ให้กู้",null,"FF0000"));

        body.AppendChild(NormalParagraphWithTabs("ข้อ 3. ดอกเบี้ย"));

        body.AppendChild(NormalParagraphWithTabs("3.1 การกู้ยืมเงินตามสัญญากู้เงินนี้ ไม่มีดอกเบี้ยเงินกู้"));
        body.AppendChild(NormalParagraphWithTabs("3.2 กรณีที่ผู้กู้ผิดเงื่อนไขการผ่อนชำระหนี้ และ/หรือไม่สามารถชำระหนี้เงินต้นคืนให้แก่ผู้ให้กู้ได้ครบถ้วนเมื่อครบกำหนดตามสัญญา" +
         "ผู้กู้และผู้ให้กู้ตกลงกันให้เป็นสิทธิของผู้ให้กู้ที่จะปรับอัตราดอกเบี้ยระหว่างผิดนัดการชำระหนี้ได้ในอัตราร้อยละ 15 ต่อปีโดยไม่ต้องบอกกล่าวผู้กู้" +
         "และ/หรือ ดำเนินการปรับโครงสร้างหนี้ให้แก่ผู้กู้ได้โดยผู้ให้กู้มีสิทธิที่จะคิดดอกเบี้ยจากผู้กู้ได้ในอัตราร้อยละ 15" +
         "ต่อปีจนกว่าจะชำระหนี้ให้แก่ผู้ให้กู้จนเสร็จสิ้น ตลอดจนดำเนินการใดๆ ได้ตามขอบเขตของประมวลกฎหมายแพ่งและพาณิชย์"));
        body.AppendChild(NormalParagraphWithTabs("ข้อ 4. การชำระคืนเงินต้นหรือชำระหนี้อื่นใด ให้แก่ผู้ให้กู้"));
        body.AppendChild(NormalParagraphWithTabs("4.1 ผู้กู้ตกลงผ่อนชำระเงินต้นคืนให้แก่ผู้ให้กู้เป็นรายเดือนไม่น้อยกว่าเดือนละ..................... บาท" +
         " (..............................................................)" +
         " โดยชำระภายในวันที่.............ของทุกเดือน เริ่มตั้งแต่เดือน..........................พ.ศ. ........ เป็นต้นไป"));

        body.AppendChild(NormalParagraphWithTabs("4.2 การชำระเงินตามข้อ 4.1  ผู้กู้ตกลงจะนำเงินเข้าบัญชีเงินฝากของผู้กู้ที่เปิดบัญชีไว้กับธนาคารกรุงไทย จำกัด (มหาชน) ตาม" +
         "ข้อ 2. โดยผู้กู้ยินยอมให้ ธนาคารพัฒนาวิสาหกิจขนาดกลางและขนาดย่อมแห่งประเทศไทยซึ่งดำเนินการแทนผู้ให้กู้ ในการแจ้งธนาคารเจ้าของบัญชีตาม" +
         "ข้อ 2. ให้หักเงินจากบัญชีของผู้กู้ดังกล่าวแล้วเพื่อชำระคืนเงินกู้แก่ผู้ให้กู้ในแต่ละงวดเดือน พร้อมทำการโอนเงินที่หักจากบัญชีของผู้กู้เพื่อนำเข้าบัญชีของผู้ให้กู้ที่เปิดบัญชี ไว้กับธนาคารกรุงไทย จำกัด (มหาชน)" +
         "สาขา............................................  บัญชีออมทรัพย์  ชื่อบัญชี" +
         "โครงการเงินทุนพลิกฟื้นวิสาหกิจขนาดย่อม เลขที่บัญชี............................................ เพื่อชำระหนี้คืนเงินกู้แก่ผู้ให้กู้ตามข้อตกลงในแต่ละงวดเดือน" +
         " เมื่อผู้ให้กู้ได้รับชำระเงินกู้คืนในแต่ละงวดแล้วจะออกใบเสร็จรับเงินให้แก่ผู้กู้ไว้เป็นหลักฐานต่อไป โดยผู้กู้ตกลงยินยอมให้หักเงินค่าธรรมเนียม\r\nในการโอนเงินชำระหนี้เงินกู้หรือค่าธรรมเนียมใด ๆ" +
         "ที่ธนาคารเจ้าของบัญชีเรียกเก็บในการโอนเงินจากบัญชีของผู้กู้ไปยังบัญชีเงินฝากของผู้ให้กู้ตามข้อ 4.2 ข้างต้นด้วย"));

        body.AppendChild(NormalParagraphWithTabs("ข้อ 5.การผิดสัญญา"));
        body.AppendChild(NormalParagraphWithTabs("5.1 ในกรณีต่อไปนี้ให้ถือว่าผู้กู้ผิดสัญญา ให้ผู้ให้กู้มีสิทธิบอกเลิกสัญญาได้"));
        body.AppendChild(NormalParagraphWithTabs("5.1.1 ผู้กู้ไม่ปฏิบัติตามสัญญาฉบับนี้ไม่ว่าข้อหนึ่งข้อใด"));
        body.AppendChild(NormalParagraphWithTabs("5.1.2 ผู้กู้ผิดนัดชำระคืนต้นเงินไม่ว่างวดหนึ่งงวดใดก็ตาม หรือเงินจำนวนอื่นใดที่ต้องชำระตามสัญญาฉบับนี้"));
        body.AppendChild(NormalParagraphWithTabs("5.1.3 ผู้กู้ให้ข้อเท็จจริง ข่าวสาร ข้อความหรือเอกสารอันเป็นเท็จ หรือปกปิด ข้อเท็จจริงซึ่งควรจะแจ้งให้ผู้ให้กู้ทราบ"));
        body.AppendChild(NormalParagraphWithTabs("5.1.4 ผู้กู้ไม่ปฏิบัติตามโครงการเงินทุนพลิกฟื้นวิสาหกิจขนาดย่อม ตามเอกสารแนบท้ายสัญญานี้"));
        body.AppendChild(NormalParagraphWithTabs("5.2 เมื่อผู้กู้ผิดสัญญาแล้วแม้ข้อหนึ่งข้อใด หรือผู้กู้ไม่ชำระหนี้ให้ถูกต้องครบถ้วนตามที่กำหนดในสัญญานี้ไม่ว่าข้อหนึ่งข้อใด หรือผิดนัดชำระหนี้งวดใด ๆให้ถือว่าเป็นการผิดนัดทั้งหมด บรรดาหนี้สินทั้งหลายที่ยังต้องชำระ\r\nอยู่ตามสัญญานี้ ไม่ว่าจะถึงกำหนดชำระแล้วหรือไม่ ให้ถือว่าเป็นอันถึงกำหนดชำระทั้งหมดทันที ผู้กู้ยินยอมให้ผู้ให้\r\nกู้คิดดอกเบี้ยจากเงินต้นที่ค้างชำระในอัตราร้อยละ 15.00 ต่อปี นับตั้งแต่วันที่ผู้กู้ตกเป็นผู้ผิดนัดตามสัญญานี้ จนกว่าจะชำระหนี้ทั้งหมดเสร็จสิ้น  พร้อมด้วยค่าเสียหายและค่าใช้จ่ายทั้งหลายอันเนื่องจากการผิดนัดชำระหนี้ของผู้กู้ รวมทั้งค่าใช้จ่าย\r\nในการเตือน เรียกร้อง ทวงถาม ดำเนินคดีและการบังคับชำระหนี้จนเต็มจำนวน\r\n"));
        
        body.AppendChild(NormalParagraphWithTabs("ข้อ 6. การเปิดเผยข้อมูล"));
        body.AppendChild(NormalParagraphWithTabs("ในการวิเคราะห์ข้อมูลเพื่อประกอบการพิจารณาให้สินเชื่อ การแก้ไขหนี้ หรือการปรับปรุงโครงสร้างหนี้ของผู้ให้กู้แก่ผู้กู้นั้น ผู้กู้ตกลงยินยอมให้ผู้ให้กู้ตรวจสอบและใช้ข้อมูลเกี่ยวกับการเงิน ประวัติและภาระหนี้  ที่ผู้กู้มีอยู่กับสถาบันการเงิน และนิติบุคคลอื่น รวมทั้งข้อมูลเครดิตของผู้กู้ที่ได้ถูกรวบรวมไว้ที่ บริษัท ข้อมูลเครดิตแห่งชาติ จำกัด  หรือบริษัทข้อมูลเครดิตใด ๆ ตามพระราชบัญญัติการประกอบธุรกิจข้อมูลเครดิต ตลอดจนการตรวจสอบการล้มละลายและหรือ \r\nการบังคับคดีขายทอดตลาดของผู้กู้ได้ โดยไม่ต้องคำนึงว่าผู้กู้จะได้รับอนุมัติสินเชื่อ ไม่ว่าจะเป็นการให้วงเงินสินเชื่อ การแก้ไขหนี้ หรือการปรับปรุงโครงสร้างหนี้จากผู้ให้กู้หรือไม่ก็ตาม\r\n"));
       
        body.AppendChild(NormalParagraphWithTabs("ข้อ 7. อื่นๆ"));
        body.AppendChild(NormalParagraphWithTabs("7.1 ในระหว่างและตลอดระยะเวลาการกู้เงินตามสัญญานี้ ผู้กู้ยินยอมให้ผู้ให้กู้ หรือตัวแทนผู้ให้กู้เข้าไปตรวจสอบกิจการ ตลอดจนเอกสารหลักฐานทางบัญชีของกิจการ สรรพสมุดและเอกสารอื่นๆ ของผู้กู้ได้ตลอด"));
        body.AppendChild(NormalParagraphWithTabs("7.2 คู่สัญญาตกลงให้ถือเอาเอกสารที่แนบท้ายสัญญานี้ บันทึกข้อตกลง และบรรดาข้อสัญญาต่างๆ เป็นส่วนหนึ่งของสัญญานี้ที่มีผลผูกพันให้ผู้กู้จะต้องปฏิบัติตาม" +
         " ซึ่งเอกสารแนบท้ายนี้อาจจะทำเพิ่มเติมในภายหลังจากวันทำสัญญานี้ โดยให้ถือเป็นส่วนหนึ่งของสัญญานี้เช่นกัน และหากเอกสารแนบท้ายสัญญาขัดหรือแย้งกันผู้กู้ตกลงปฏิบัติตามคำวินิจฉัยของผู้ให้กู้"));
        body.AppendChild(NormalParagraphWithTabs("7.3 บรรดาหนังสือ จดหมาย คำบอกกล่าวใดๆ เช่น การทวงถาม การบอกเลิกสัญญา ของผู้ให้กู้ที่ส่งไปยังสถานที่ที่ระบุไว้ว่าเป็นที่อยู่ของผู้กู้ข้างต้น" +
         "หรือสถานที่อยู่ที่ผู้กู้แจ้งเปลี่ยนแปลง โดยส่งเองหรือส่งทางไปรษณีย์ลงทะเบียน หรือไม่ลงทะเบียนไม่ว่าจะมีผู้รับไว้หรือไม่มีผู้ใดยอมรับไว้" +
         "หรือส่งไม่ได้เพราะผู้กู้ย้ายสถานที่อยู่ไปโดยมิได้แจ้งให้ผู้ให้กู้ทราบให้ไว้นั้นหาไม่พบ หรือถูกรื้อถอนทำลายทุกๆ กรณีดังกล่าวให้ถือว่าผู้กู้ได้รับโดยชอบแล้ว"));
        body.AppendChild(NormalParagraphWithTabs("7.4 การสละสิทธิ์ตามสัญญานี้ ในคราวหนึ่งคราวใดของผู้ให้กู้ หรือการที่ผู้ให้กู้มิได้ใช้สิทธิ์ที่มีอยู่ ไม่ถือเป็นการสละสิทธิ์ของผู้ให้กู้ในคราวต่อไปและไม่มีผลกระทบต่อการใช้สิทธิของผู้ให้กู้ในคราวต่อไป"));
        body.AppendChild(NormalParagraphWithTabs("7.5 หากข้อกำหนด และ/หรือเงื่อนไขข้อใดข้อหนึ่งของสัญญานี้ตกเป็นโมฆะ หรือใช้บังคับไม่ได้ตามกฎหมาย ให้ข้อกำหนดและเงื่อนไขอื่น ๆ ยังคงมีผลใช้บังคับได้ต่อไปได้ โดยแยกต่างหากจากส่วนที่เป็นโมฆะหรือไม่สมบูรณ์นั้น"));
        body.AppendChild(NormalParagraphWithTabs("ผู้กู้ได้ตรวจ อ่าน และเข้าใจข้อความในสัญญานี้โดยละเอียดโดยตลอดแล้ว เห็นว่าถูกต้องตามเจตนาทุกประการ จึงได้ลงลายมือชื่อพร้อมประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยาน ณ วัน เดือน ปี ที่ระบุไว้ข้างต้น"));






                body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

                body.AppendChild(CenteredParagraph("ลงชื่อ........................................................................ผู้กู้"));
        body.AppendChild(CenteredParagraph("(................................................................................)"));
        body.AppendChild(CenteredParagraph("ลงชื่อ.......................................................................คู่สมรสให้ความยินยอม"));
        body.AppendChild(CenteredParagraph("(................................................................................)"));
        body.AppendChild(CenteredParagraph("ลงชื่อ.......................................................................คู่สมรสให้ความยินยอม"));
        body.AppendChild(CenteredParagraph("(...............................................................................)"));
        body.AppendChild(CenteredParagraph("ลงชื่อ......................................................................พยาน"));
        body.AppendChild(CenteredParagraph("(...............................................................................)"));
        body.AppendChild(CenteredParagraph("ลงชื่อ.....................................................................พยาน"));
        body.AppendChild(CenteredParagraph("(...............................................................................)"));

        // next page
        body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

        body.AppendChild(CenteredParagraph("คำรับรองสถานภาพการสมรส"));
        body.AppendChild(NormalParagraphWithTabs("ข้าพเจ้า………………………………………………………………………………………………………………………………………….\r\nขอรับรองว่าสถานภาพการสมรสของข้าพเจ้าปัจจุบันมีสถานะ\r\n"));
        body.AppendChild(NormalParagraphWithTabs("“ข้าพเจ้าขอรับรองว่าสถานภาพการสมรสที่แจ้งในหนังสือฉบับนี้เป็นความจริงทุกประการหากไม่เป็นความจริงแล้ว ความเสียหายใด ๆ ที่จะเกิดกับผู้ให้การอุดหนุน ข้าพเจ้ายินยอมรับผิดชดใช้ให้แก่ผู้ให้การอุดหนุนทั้งสิ้น”"));
        body.AppendChild(CenteredParagraph("ลงชื่อ.............................................................รับรอง"));
        body.AppendChild(CenteredParagraph("(............................................................)"));
        body.AppendChild(CenteredParagraph("ลงชื่อ....................................................พยาน          ลงชื่อ ........................................................พยาน"));
        body.AppendChild(CenteredParagraph("(............................................................)                 (.........................................................)"));

        body.AppendChild(EmptyParagraph());
        body.AppendChild(RightParagraph("........................................................./ผู้พิมพ์"));
        body.AppendChild(RightParagraph("........................................................./ผู้ตรวจ"));


                // --- Add header for first page (empty) ---
                var firstHeaderPart = mainPart.AddNewPart<HeaderPart>();
                string firstHeaderPartId = mainPart.GetIdOfPart(firstHeaderPart);
                firstHeaderPart.Header = new Header(
                    new Paragraph() // Empty paragraph, so no page number on first page
                );

                // --- Add header for other pages (centered page number) ---
                var headerPart = mainPart.AddNewPart<HeaderPart>();
                string headerPartId = mainPart.GetIdOfPart(headerPart);
                headerPart.Header = new Header(
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
                    new HeaderReference() { Type = HeaderFooterValues.First, Id = firstHeaderPartId },
                    new HeaderReference() { Type = HeaderFooterValues.Default, Id = headerPartId },
                    new PageSize() { Width = 11906, Height = 16838 }, // A4 size
                    new PageMargin() { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720, Gutter = 0 },
                    new TitlePage() // This enables different first page header/footer
                );
                body.AppendChild(sectionProps);
            }
            stream.Position = 0;
            return File(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สสว.สัญญาเงินกู้ยืมโครงการพลิกฟื้นวิสาห.docx");
        }
        #endregion
        private static Paragraph CenteredBoldColoredParagraph(string text, string hexColor) =>
            new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                        new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" },
                        new Bold(),
                        new Color { Val = hexColor }
                    ),
                    new Text(text)
                )
            );

        #endregion
    }
}