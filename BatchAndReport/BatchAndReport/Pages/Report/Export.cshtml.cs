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
        public ExportModel(SmeDAO smeDao, IWordEContract_AllowanceService wordEContract_AllowanceService)
        {
            _smeDao = smeDao;
            this._wordEContract_AllowanceService = wordEContract_AllowanceService;
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

        private static Paragraph JustifiedParagraph(string text,string fontSize ="28") =>
    new Paragraph(
        new ParagraphProperties(new Justification { Val = JustificationValues.Both }),
        new Run(
            new RunProperties(
                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = fontSize }
            ),
            new Text(text)
        )
    );
        // Helper: Paragraph with 2 tab spaces at the start of the first line
        private static Paragraph NormalParagraphWith_1Tabs(string text, JustificationValues? align = null, string fontZise = "28")
        {
            if (fontZise == null)
            {
                fontZise = "28";
            }
            var paragraph = new Paragraph();

            // Paragraph properties (alignment and tab stops)
            var props = new ParagraphProperties();
            if (align != null)
                props.Append(new Justification { Val = align.Value });

            // Add two tab stops (every 720 = 0.5 inch, adjust as needed)
            var tabs = new Tabs(
                new TabStop { Val = TabStopValues.Left, Position = 720 }               
            );
            props.Append(tabs);
            paragraph.Append(props);

            // Add two tab characters at the start
            var run = new Run(
                new RunProperties(
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = fontZise }
                ),
                new TabChar(),
             
                new Text(text)
            );
            paragraph.Append(run);

            return paragraph;
        }
        private static Paragraph NormalParagraphWith_2Tabs(string text, JustificationValues? align = null, string fontZise = "28")
        {
            if (fontZise == null)
            {
                fontZise = "28";
            }
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
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = fontZise }
                ),
                new TabChar(),
                new TabChar(),
                new Text(text)
            );
            paragraph.Append(run);

            return paragraph;
        }
        private static Paragraph NormalParagraphWith_3Tabs(string text, JustificationValues? align = null, string fontZise = "28")
        {
            if (fontZise == null)
            {
                fontZise = "28";
            }
            var paragraph = new Paragraph();

            // Paragraph properties (alignment and tab stops)
            var props = new ParagraphProperties();
            if (align != null)
                props.Append(new Justification { Val = align.Value });

            // Add three tab stops (every 720 = 0.5 inch, adjust as needed)
            var tabs = new Tabs(
                new TabStop { Val = TabStopValues.Left, Position = 720 },
                new TabStop { Val = TabStopValues.Left, Position = 1440 },
                new TabStop { Val = TabStopValues.Left, Position = 2160 }
            );
            props.Append(tabs);
            paragraph.Append(props);

            // Add three tab characters at the start
            var run = new Run(
                new RunProperties(
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = fontZise }
                ),
                new TabChar(),
                new TabChar(),
                new TabChar(),
                new Text(text)
            );
            paragraph.Append(run);

            return paragraph;
        }
        private static Paragraph NormalParagraphWith_2TabsColor(string text, JustificationValues? align = null, string hexColor = null)
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

        private static Paragraph CenteredBoldColoredParagraph(string text, string hexColor,string fonsize="28") =>
          new Paragraph(
              new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
              new Run(
                  new RunProperties(
                      new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                      new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = fonsize },
                      new Bold(),
                      new Color { Val = hexColor }
                  ),
                  new Text(text)
              )
          );

        private static void AddHeaderWithPageNumber(MainDocumentPart mainPart, Body body)
        {
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
                        new FieldCode(" PAGE") { Space = SpaceProcessingModeValues.Preserve }
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

        #region สสว. สัญญารับเงินอุดหนุน
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
                    var element = CreateImage(mainPart.GetIdOfPart(imagePart), 240, 80);
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
                body.AppendChild(CenteredBoldParagraph("ที่ศูนย์ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม"));
                body.AppendChild(CenteredBoldParagraph("วันที่...................................................."));

                // 3. Fillable lines (using underlines)
                //  body.AppendChild(EmptyParagraph());
                body.AppendChild(NormalParagraphWith_2Tabs("ข้าพเจ้า ...................................................................................................................."));
                body.AppendChild(JustifiedParagraph("อายุ ......... ปี สัญชาติ .................. สำนักงาน/บ้านตั้งอยู่เลขที่.................. อาคาร..........................................."));
                body.AppendChild(JustifiedParagraph("หมู่ที่...........ตรอก/ซอย..........................ถนน...........................ตำบล/แขวง.................. ..................."));
                body.AppendChild(JustifiedParagraph("เขต/อำเภอ................... จังหวัด.................ทะเบียนนิติบุคคลเลขที่/เลขประจำตัวประชาชนที่............................................"));
                body.AppendChild(JustifiedParagraph("จดทะเบียนเป็นนิติบุคคลเมื่อวันที่ .........................................."));

                // 4. Main body (sample)
                //   body.AppendChild(EmptyParagraph());
                body.AppendChild(NormalParagraph("ซึ่งต่อไปนี้จะเรียกบุคคลผู้มีนามตามที่ระบุข้างต้นทั้งหมดว่า \"ผู้รับการอุดหนุน\" ได้ทำสัญญาฉบับนี้ให้ไว้แก่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งต่อไปนี้จะเรียกว่า \"ผู้ให้การอุดหนุน\" โดยมีสาระสำคัญดังนี้"));
                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ1. ผู้รับการอุดหนุนได้ขอรับความช่วยเหลือผ่านการอุดหนุนตามมาตรการฟื้นฟูกิจการวิสาหกิจ ขนาดกลางและขนาดย่อมจากผู้ให้การอุดหนุนเป็นจำนวนเงิน ..................... บาท (...........................) ปลอดการชำระเงินต้น ................. เดือน โดยไม่มีดอกเบี้ย แต่มีภาระต้องชำระคืนเงินต้น "));
                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ2. ผู้ให้การอุดหนุนจะให้ความช่วยเหลือด้วยการให้เงินอุดหนุนแก่ผู้รับการอุดหนุน ด้วยการนำเงินหรือโอนเงินเข้าบัญชีธนาคารกรุงไทย จำกัด (มหาชน) สาขา ..................................... เลขที่บัญชี ................................................. ชื่อบัญชี .......................... ซึ่งเป็นบัญชีของผู้รับการอุดหนุน จำนวนเงิน ........................... บาท (..............................) และให้ถือว่าผู้รับการอุดหนุนได้รับเงินอุดหนุนตามสัญญานี้ไปจากผู้ให้การอุดหนุนแล้ว ในวันที่เงินเข้าบัญชีของผู้รับการอุดหนุนดังกล่าว"));
                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ3. ห้ามผู้รับการอุดหนุนนำเงินอุดหนุนไปชำระหนี้เดิมที่มีอยู่ก่อนทำสัญญานี้"));
                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ4. ผู้รับการอุดหนุนยินยอมให้ธนาคารกรุงไทย จำกัด (มหาชน) ซึ่งกระทำการแทนผู้ให้การอุดหนุน หักเงินอุดหนุนที่จะได้จากผู้ให้การอุดหนุนเป็นค่าใช้จ่ายหรือค่าธรรมเนียมในการโอนเงินเข้าบัญชีของผู้รับการอุดหนุน ซึ่งธนาคารกรุงไทย จำกัด (มหาชน) เรียกเก็บตามระเบียบของธนาคารได้ โดยไม่ต้องบอกกล่าวหรือแจ้งให้ผู้รับการอุดหนุนทราบล่วงหน้า และให้ถือว่าผู้รับการอุดหนุนได้รับเงินตามจำนวนที่เบิกไปครบถ้วนแล้ว"));
                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ5. ผู้รับการอุดหนุนตกลงผ่อนชำระเงินต้นคืนให้แก่ผู้ให้การอุดหนุนเป็นรายเดือน (งวด) ๆ ละ ไม่น้อยกว่า ....................... บาท (.....................................) ด้วยการโอนเข้าบัญชีตามที่ระบุไว้ในข้อ 2 โดยชำระเงินต้นงวดแรกในเดือนที่ ....................... นับถัดจากวันที่ได้รับเงินอุดหนุน และงวดถัดไปทุกวันที่ .................. ของเดือนจนกว่าจะชำระเสร็จสิ้น  แต่ทั้งนี้จะต้องชำระให้เสร็จสิ้นไม่เกินกว่า .............. ปี (...........) นับแต่วันที่ได้รับเงินอุดหนุน"));
                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ6. การชำระเงินคืนตาม"));
                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ5. ผู้รับการอุดหนุนตกลงจะนำเงินเข้าบัญชีเงินฝากของผู้รับการอุดหนุน ที่เปิดบัญชีไว้กับธนาคารกรุงไทย จำกัด (มหาชน) ตามข้อ 2 โดยผู้รับการอุดหนุนยินยอมให้ ธนาคารกรุงไทย จำกัด (มหาชน) ซึ่งดำเนินการแทนผู้ให้การอุดหนุน หักเงินจากบัญชีของผู้รับการอุดหนุนดังกล่าวเพื่อชำระคืนเงินอุดหนุนแก่ผู้ให้การอุดหนุน"));
                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ6. การชำระเงินคืนตามข้อ 5 ผู้รับการอุดหนุนตกลงจะนำเงินเข้าบัญชีเงินฝากของผู้รับการอุดหนุนที่เปิดบัญชีไว้กับธนาคารกรุงไทย จำกัด (มหาชน) ตามข้อ 2 โดยผู้รับการอุดหนุนยินยอมให้ ธนาคารกรุงไทย จำกัด (มหาชน) ซึ่งดำเนินการแทนผู้ให้การอุดหนุน หักเงินจากบัญชีของผู้รับการอุดหนุนดังกล่าวเพื่อชำระคืนเงินอุดหนุนแก่ผู้ให้การอุดหนุน ในแต่ละงวดเดือน พร้อมทำการโอนเงินที่พักจากบัญชีของผู้รับการอุดหนุนมอบเข้าบัญชีของผู้ให้การอุดหนุนที่เปิดบัญชีไว้กับธนาคารกรุงไทย จำกัด (มหาชน) สาขา .....องค์การตลาดเพื่อเกษตรกร (จตุจักร)..... บัญชีออมทรัพย์ ชื่อบัญชีสำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เลขที่บัญชี ....035-1-52709-5.....เพื่อชำระหนี้คืนเงินอุดหนุนแก่ผู้ให้การอุดหนุนตามข้อตกลงในแต่ละงวดเดือน"));
                body.AppendChild(NormalParagraphWith_2Tabs("ไม่ว่าผู้รับการอุดหนุนจะได้จัดทำหนังสือยินยอมให้ธนาคารกรุงไทย จำกัด (มหาชน) หักบัญชีเงินฝาก ตามวรรคหนึ่งหรือไม่ก็ตาม โดยสัญญนี้ผู้รับการอุดหนุนให้ถือว่าเป็นการทำหนังสือยินยอมให้ธนาคารกรุงไทย จำกัด (มหาชน) หักบัญชีเงินฝากตามวรรคหนึ่งด้วย"));
                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ7. ผู้รับการอุดหนุนตกลงยินยอมให้ธนาคารกรุงไทย จจำกัด (มหาชน) ซึ่งกระทำการแทนผู้ให้การอุดหนุนหักเงินที่ผู้รับการอุดหนุนได้โอนเข้าบัญชีตามข้อ 2 เพื่อชำระคืนเป็นค่าใช้จ่ายหรือค่าธรรมเนียมในการโอนเงิน"));
                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ8. ในระหว่างและตลอดระยะเวลาการตามสัญญาฉบับนี้ ผู้รับการอุดหนุนจะต้องรายงานผลการประกอบกิจการมายังผู้ให้การอุดหนุนหรือศูนย์ให้บริหร SMEs ครบวงจร ในจังหวัดที่ผู้รับการอุดหนุนมีภูมิลำเนาอยู่หรือพื้นที่ใกล้เคียงหรือหน่วยงานอื่นใดที่ผู้ให้การอุดหนุนมอบหมาย ตามหลักเกณฑ์และวิธีการที่ผู้ให้การอุดหนุนกำหนด ไม่น้อยกว่าเดือนละหนึ่งครั้ง"));
                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ9. กรณีต่อไปนี้ให้ถือว่าผู้รับการอุดหนุนปฏิบัติผิดสัญญา"));
                body.AppendChild(NormalParagraphWith_2Tabs("9.1 ผู้รับการอุดหนุนผิดนัดชำระคืนเงินอุดหนุนไม่ว่างวดหนึ่งวดใดก็ตาม หรือไม่ชำระคืนเงินอุดหนุนภายในกำหนดระยะเวลาที่กำหนดในสัญญานี้ หรือเงินจำนวนอื่นใดที่ต้องชำระตามสัญญาฉบับนี้"));
                body.AppendChild(NormalParagraphWith_2Tabs("9.2 ผู้รับการอุดหนุนใช้เงินอุดหนุนผิดไปจากเงื่อนไขตามสัญญา หรือผิดสัญญาแม้ข้อใดข้อหนึ่ง หรือไม่รายงานการดำเนินธุรกิจให้ผู้ให้การอุดหนุนทราบตามข้อ 8 หรือตรวจสอบในภายหลังแล้วพบว่ามีการแจ้งคุณสมบัติ หรือส่งเอกสารเป็นเท็จแก่ผู้ให้การอุดหนุน ๆ มีสิทธิบอกเลิกสัญญาได้"));
                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ10. อื่นๆ"));
                body.AppendChild(NormalParagraphWith_2Tabs("10.1 ในระหว่างและตลอดระยะเวลาตามสัญญานี้ ผู้รับการอุดหนุนยินยอมให้ผู้ให้การอุดหนุน หรือตัวแทนผู้ให้การอุดหนุนเข้าไปตรวจสอบติดตามการดำเนินธุรกิจ ตลอดจนเอกสารหลักฐานทางบัญชีของกิจการ สรรพเอกสารอื่น ๆ ของผู้รับการอุดหนุนได้ตลอด"));
                body.AppendChild(NormalParagraphWith_2Tabs("10.2 คู่สัญญาตกลงให้ถือเอาเอกสารที่แนบท้ายสัญญานี้ บันทึกข้อตกลง และบรรดาข้อสัญญาต่าง ๆ  เป็นส่วนหนึ่งของสัญญานี้ที่มีผลผูกพันให้ผู้รับการอุดหนุนจะต้องปฏิบัติตาม ซึ่งเอกสารแนบท้ายนี้อาจจะทำเพิ่มเติมในภายหลังจากวันทำสัญญานี้ โดยให้ถือเป็นส่วนหนึ่งของสัญญานี้เช่นกัน และหากเอกสารแนบท้ายสัญญาขัดหรือแย้งกันผู้รับการอุดหนุนตกลงปฏิบัติตามคำวินิจฉัยของผู้ให้การอุดหนุน"));
                body.AppendChild(NormalParagraphWith_2Tabs("10.3 บรรดาหนังสือ จดหมาย คำบอกกล่าวใด ๆ เช่น การทวงถาม การบอกเลิกสัญญา ของผู้ให้การอุดหนุนหรือผู้ที่ได้รับมอบหมายส่งไปยังสถานที่ที่ระบุไว้เป็นที่อยู่ของผู้รับการอุดหนุนข้างต้น หรือสถานที่อยู่ที่ผู้รับการอุดหนุนแจ้งเปลี่ยนแปลง โดยส่งเองหรือส่งทางไปรษณีย์ลงทะเบียน หรือไม่ลงทะเบียน ไม่ว่าจะมีผู้รับไว้ หรือไม่มีผู้ใดยอมรับไว้ หรือส่งไม่ได้เพราะผู้รับการอุดหนุนย้ายสถานที่อยู่ไปโดยมิได้แจ้งให้ผู้ให้การอุดหนุนทราบหรือหาไม่พบ หรือถูกรื้อถอนทำลายทุก ๆ กรณีดังกล่าวให้ถือว่าผู้รับการอุดหนุนได้รับโดยชอบแล้ว"));
                body.AppendChild(NormalParagraphWith_2Tabs("10.4 การสละสิทธิ์ตามสัญญานี้ ในคราวหนึ่งคราวใดของผู้ให้การอุดหนุน หรือการที่ผู้ให้การอุดหนุนมิได้ ใช้สิทธิ์ที่มีอยู่ ไม่ถือเป็นการสละสิทธิ์ของผู้ให้การอุดหนุนในคราวต่อไปและไม่มีผลกระทบต่อการใช้สิทธิของผู้ให้การอุดหนุน ในคราวต่อไป"));
                body.AppendChild(NormalParagraphWith_2Tabs("10.5 หากข้อกำหนด และ/หรือเงื่อนไขข้อใดข้อหนึ่งของสัญญานี้ตกเป็นโมฆะ หรือใช้บังคับไม่ได้ตามกฎหมาย ให้ข้อกำหนดและเงื่อนไขอื่น ๆ ยังคงมีผลใช้บังคับได้ต่อไปได้ โดยแยกต่างหากจากส่วนที่เป็นโมฆะหรือไม่สมบูรณ์นั้น"));
                body.AppendChild(NormalParagraphWith_2Tabs("สัญญานี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน คู่สัญญาทั้งสองฝ่ายได้ตรวจ อ่าน และเข้าใจข้อความในสัญญานี้โดยละเอียดแล้ว เห็นว่าถูกต้องตามเจตนาทุกประการ จึงได้ลงลายมือชื่อพร้อมประทับตรา (ถ้ามี) ไว้เป็นสำคัญ ต่อหน้าพยาน ณ วัน เดือน ปี ที่ระบุไว้ข้างต้น"));

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
                body.AppendChild(NormalParagraphWith_2Tabs("ข้าพเจ้า………………………………………………………………………………………………………………………………………….\r\nขอรับรองว่าสถานภาพการสมรสของข้าพเจ้าปัจจุบันมีสถานะ\r\n"));
                body.AppendChild(NormalParagraphWith_2Tabs("“ข้าพเจ้าขอรับรองว่าสถานภาพการสมรสที่แจ้งในหนังสือฉบับนี้เป็นความจริงทุกประการหากไม่เป็นความจริงแล้ว ความเสียหายใด ๆ ที่จะเกิดกับผู้ให้การอุดหนุน ข้าพเจ้ายินยอมรับผิดชดใช้ให้แก่ผู้ให้การอุดหนุนทั้งสิ้น”"));
                body.AppendChild(CenteredParagraph("ลงชื่อ.............................................................รับรอง"));
                body.AppendChild(CenteredParagraph("(............................................................)"));
                body.AppendChild(CenteredParagraph("ลงชื่อ....................................................พยาน                    ลงชื่อ ........................................................พยาน"));
                body.AppendChild(CenteredParagraph("(............................................................)                                 (.........................................................)"));

                body.AppendChild(EmptyParagraph());
                body.AppendChild(RightParagraph("........................................................./ผู้พิมพ์"));
                body.AppendChild(RightParagraph("........................................................./ผู้ตรวจ"));


                AddHeaderWithPageNumber(mainPart, body);
            }
            stream.Position = 0;
            return File(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สสว.สัญญารับเงินอุดหนุน.docx");
        }
        // Helper for colored, bold, centered paragraph
        #endregion

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
                    var element = CreateImage(mainPart.GetIdOfPart(imagePart), 240, 80);
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
                body.AppendChild(CenteredBoldColoredParagraph("สัญญากู้ยืมเงิน", "FF0000")); // Blue
                body.AppendChild(CenteredBoldColoredParagraph("โครงการเงินทุนพลิกฟื้นวิสาหกิจขนาดย่อม", "FF0000")); // Red
                body.AppendChild(RightParagraph("ทำที่ ธนาคารพัฒนาวิสาหกิจขนาดกลางและขนาดย่อมแห่งประเทศไทย"));
                body.AppendChild(RightParagraph("สำนักงานใหญ่/สาขา.........................................................."));

                // 3. Fillable lines (using underlines)
                // body.AppendChild(EmptyParagraph());
                body.AppendChild(NormalParagraphWith_2Tabs("ข้าพเจ้า ...................................................................................................................."));
                body.AppendChild(JustifiedParagraph("อายุ ......... ปี สัญชาติ .................. สำนักงาน/บ้านตั้งอยู่เลขที่.................. อาคาร..........................................."));
                body.AppendChild(JustifiedParagraph("หมู่ที่...........ตรอก/ซอย..........................ถนน...........................ตำบล/แขวง.................. ..................."));
                body.AppendChild(JustifiedParagraph("เขต/อำเภอ................... จังหวัด.................ทะเบียนนิติบุคคลเลขที่/เลขประจำตัวประชาชนที่............................................"));
                body.AppendChild(JustifiedParagraph("จดทะเบียนเป็นนิติบุคคลเมื่อวันที่..................................ซึ่งต่อไปนี้จะเรียกบุคคลผู้มีนามตามที่ระบุข้างต้นทั้งหมดว่า \"ผู้กู้\"ได้ทำสัญญาฉบับนี้ให้ไว้แก่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม" +
                 "ซึ่งต่อไปนี้จะเรียกว่า \"ผู้ให้กู้\" โดยมีสาระสำคัญดังนี้"));

                // 4. Main body (sample)
                //  body.AppendChild(EmptyParagraph());
                body.AppendChild(NormalParagraph("ซึ่งต่อไปนี้จะเรียกบุคคลผู้มีนามตามที่ระบุข้างต้นทั้งหมดว่า \"ผู้รับการอุดหนุน\" ได้ทำสัญญาฉบับนี้ให้ไว้แก่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม " +
                    "ซึ่งต่อไปนี้จะเรียกว่า \"ผู้ให้การอุดหนุน\" โดยมีสาระสำคัญดังนี้", null, "28"));
                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 1. วัตถุประสงค์และวงเงินกู้"));
                body.AppendChild(NormalParagraphWith_2Tabs("โดยผู้กู้ได้กู้เงินจากผู้ให้กู้เป็นจำนวนเงิน.....................................บาท(....................................................)เ" +
                 "พื่อนำไปใช้จ่ายเป็นเงินทุนหมุนเวียน"));
                body.AppendChild(NormalParagraphWith_2TabsColor("โดยไม่นำเงินที่กู้ยืมไปชำระหนี้ที่มีอยู่ก่อนยื่นคำขอกู้ยืมเงิน", null, "FFF0000"));

                body.AppendChild(NormalParagraphWith_2Tabs("กำหนดชำระเงินกู้เสร็จสิ้นภายใน...........ปี...........เดือน  โดยมีระยะเวลาปลอดเงินต้น................เดือน"));
                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 2. การเบิกจ่ายเงินกู้"));
                body.AppendChild(NormalParagraphWith_2Tabs("ผู้ให้กู้จะจ่ายเงินกู้แก่ผู้กู้ตามเงื่อนไขการใช้เงินกู้ในข้อ 1.และตามรายละเอียดการใช้เงินกู้ ซึ่งผู้กู้ได้แจ้งไว้ในคำขอสินเชื่อและเอกสารแนบท้ายคำขอสินเชื่อโดยถือเป็นส่วนหนึ่งของสัญญากู้เงินฉบับนี้ด้วย" +
                 "หากปรากฏว่ารายการขอเบิกเงินกู้งวดใดไม่เป็นไปตามเงื่อนไขและรายละเอียดดังกล่าว " +
                 "เป็นสิทธิของผู้ให้กู้แต่ฝ่ายเดียวที่จะพิจารณาไม่ให้เบิกเงินกู้ก็ได้"));
                body.AppendChild(NormalParagraphWith_2TabsColor("โดยผู้ให้กู้จะจ่ายเงินกู้ให้ผู้กู้ด้วยการนำเงินหรือโอนเงินเข้าบัญชีที่ ธนาคารกรุงไทย จำกัด (มหาชน)" +
                 "\r\nสาขา..............................................ชื่อบัญชี................................................................................ซึ่งเป็นบัญชีของผู้กู้" +
                 "\r\nเลขที่บัญชี...................................จำนวนเงิน...........................บาท (.............................................)" +
                 "และให้ถือว่าผู้กู้ได้รับเงินกู้ตามสัญญานี้ไปจากผู้ให้กู้แล้ว ในวันที่เงินเข้าบัญชีของผู้กู้ดังกล่าว"));
                body.AppendChild(NormalParagraphWith_2TabsColor("ทั้งนี้ ผู้กู้ยินยอมให้ผู้ให้กู้ หรือ ธนาคารพัฒนาวิสาหกิจขนาดกลางและขนาดย่อมแห่งประเทศไทย" +
                 "ซึ่งกระทำการแทนผู้ให้กู้ หักเงินจากจำนวนเงินกู้ที่ผู้กู้ขอเบิกจากผู้ให้กู้เป็นค่าวิเคราะห์โครงการ ค่าอากรแสตมป์ ค่าใช้จ่ายหรือค่าธรรมเนียมในการโอนเงินเข้าบัญชีของผู้กู้ซึ่งธนาคารกรุงไทย จำกัด (มหาชน)" +
                 " เรียกเก็บตามระเบียบของธนาคาร โดยไม่ต้องบอกกล่าวหรือแจ้งให้ผู้กู้ทราบ" +
                 "โดยให้ถือว่าผู้กู้ได้รับเงินกู้ตามจำนวนที่เบิกไปครบถ้วนแล้วและสละสิทธิ์ที่จะเรียกร้องอย่างใด ๆ" +
                 "ต่อผู้ให้กู้และหรือธนาคารพัฒนาวิสาหกิจขนาดกลางและขนาดย่อมแห่งประเทศไทย ที่ดำเนินการแทนตามที่ได้รับมอบหมายจากผู้ให้กู้", null, "FF0000"));

                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 3. ดอกเบี้ย"));

                body.AppendChild(NormalParagraphWith_2Tabs("3.1 การกู้ยืมเงินตามสัญญากู้เงินนี้ ไม่มีดอกเบี้ยเงินกู้"));
                body.AppendChild(NormalParagraphWith_2Tabs("3.2 กรณีที่ผู้กู้ผิดเงื่อนไขการผ่อนชำระหนี้ และ/หรือไม่สามารถชำระหนี้เงินต้นคืนให้แก่ผู้ให้กู้ได้ครบถ้วนเมื่อครบกำหนดตามสัญญา" +
                 "ผู้กู้และผู้ให้กู้ตกลงกันให้เป็นสิทธิของผู้ให้กู้ที่จะปรับอัตราดอกเบี้ยระหว่างผิดนัดการชำระหนี้ได้ในอัตราร้อยละ 15 ต่อปีโดยไม่ต้องบอกกล่าวผู้กู้" +
                 "และ/หรือ ดำเนินการปรับโครงสร้างหนี้ให้แก่ผู้กู้ได้โดยผู้ให้กู้มีสิทธิที่จะคิดดอกเบี้ยจากผู้กู้ได้ในอัตราร้อยละ 15" +
                 "ต่อปีจนกว่าจะชำระหนี้ให้แก่ผู้ให้กู้จนเสร็จสิ้น ตลอดจนดำเนินการใดๆ ได้ตามขอบเขตของประมวลกฎหมายแพ่งและพาณิชย์"));
                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 4. การชำระคืนเงินต้นหรือชำระหนี้อื่นใด ให้แก่ผู้ให้กู้"));
                body.AppendChild(NormalParagraphWith_2Tabs("4.1 ผู้กู้ตกลงผ่อนชำระเงินต้นคืนให้แก่ผู้ให้กู้เป็นรายเดือนไม่น้อยกว่าเดือนละ..................... บาท" +
                 " (..............................................................)" +
                 " โดยชำระภายในวันที่.............ของทุกเดือน เริ่มตั้งแต่เดือน..........................พ.ศ. ........ เป็นต้นไป"));

                body.AppendChild(NormalParagraphWith_2Tabs("4.2 การชำระเงินตามข้อ 4.1  ผู้กู้ตกลงจะนำเงินเข้าบัญชีเงินฝากของผู้กู้ที่เปิดบัญชีไว้กับธนาคารกรุงไทย จำกัด (มหาชน) ตาม" +
                 "ข้อ 2. โดยผู้กู้ยินยอมให้ ธนาคารพัฒนาวิสาหกิจขนาดกลางและขนาดย่อมแห่งประเทศไทยซึ่งดำเนินการแทนผู้ให้กู้ ในการแจ้งธนาคารเจ้าของบัญชีตาม" +
                 "ข้อ 2. ให้หักเงินจากบัญชีของผู้กู้ดังกล่าวแล้วเพื่อชำระคืนเงินกู้แก่ผู้ให้กู้ในแต่ละงวดเดือน พร้อมทำการโอนเงินที่หักจากบัญชีของผู้กู้เพื่อนำเข้าบัญชีของผู้ให้กู้ที่เปิดบัญชี ไว้กับธนาคารกรุงไทย จำกัด (มหาชน)" +
                 "สาขา............................................  บัญชีออมทรัพย์  ชื่อบัญชี" +
                 "โครงการเงินทุนพลิกฟื้นวิสาหกิจขนาดย่อม เลขที่บัญชี............................................ เพื่อชำระหนี้คืนเงินกู้แก่ผู้ให้กู้ตามข้อตกลงในแต่ละงวดเดือน" +
                 " เมื่อผู้ให้กู้ได้รับชำระเงินกู้คืนในแต่ละงวดแล้วจะออกใบเสร็จรับเงินให้แก่ผู้กู้ไว้เป็นหลักฐานต่อไป โดยผู้กู้ตกลงยินยอมให้หักเงินค่าธรรมเนียม\r\nในการโอนเงินชำระหนี้เงินกู้หรือค่าธรรมเนียมใด ๆ" +
                 "ที่ธนาคารเจ้าของบัญชีเรียกเก็บในการโอนเงินจากบัญชีของผู้กู้ไปยังบัญชีเงินฝากของผู้ให้กู้ตามข้อ 4.2 ข้างต้นด้วย"));

                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 5.การผิดสัญญา"));
                body.AppendChild(NormalParagraphWith_2Tabs("5.1 ในกรณีต่อไปนี้ให้ถือว่าผู้กู้ผิดสัญญา ให้ผู้ให้กู้มีสิทธิบอกเลิกสัญญาได้"));
                body.AppendChild(NormalParagraphWith_2Tabs("5.1.1 ผู้กู้ไม่ปฏิบัติตามสัญญาฉบับนี้ไม่ว่าข้อหนึ่งข้อใด"));
                body.AppendChild(NormalParagraphWith_2Tabs("5.1.2 ผู้กู้ผิดนัดชำระคืนต้นเงินไม่ว่างวดหนึ่งงวดใดก็ตาม หรือเงินจำนวนอื่นใดที่ต้องชำระตามสัญญาฉบับนี้"));
                body.AppendChild(NormalParagraphWith_2Tabs("5.1.3 ผู้กู้ให้ข้อเท็จจริง ข่าวสาร ข้อความหรือเอกสารอันเป็นเท็จ หรือปกปิด ข้อเท็จจริงซึ่งควรจะแจ้งให้ผู้ให้กู้ทราบ"));
                body.AppendChild(NormalParagraphWith_2Tabs("5.1.4 ผู้กู้ไม่ปฏิบัติตามโครงการเงินทุนพลิกฟื้นวิสาหกิจขนาดย่อม ตามเอกสารแนบท้ายสัญญานี้"));
                body.AppendChild(NormalParagraphWith_2Tabs("5.2 เมื่อผู้กู้ผิดสัญญาแล้วแม้ข้อหนึ่งข้อใด หรือผู้กู้ไม่ชำระหนี้ให้ถูกต้องครบถ้วนตามที่กำหนดในสัญญานี้ไม่ว่าข้อหนึ่งข้อใด หรือผิดนัดชำระหนี้งวดใด ๆให้ถือว่าเป็นการผิดนัดทั้งหมด บรรดาหนี้สินทั้งหลายที่ยังต้องชำระ\r\nอยู่ตามสัญญานี้ ไม่ว่าจะถึงกำหนดชำระแล้วหรือไม่ ให้ถือว่าเป็นอันถึงกำหนดชำระทั้งหมดทันที ผู้กู้ยินยอมให้ผู้ให้\r\nกู้คิดดอกเบี้ยจากเงินต้นที่ค้างชำระในอัตราร้อยละ 15.00 ต่อปี นับตั้งแต่วันที่ผู้กู้ตกเป็นผู้ผิดนัดตามสัญญานี้ จนกว่าจะชำระหนี้ทั้งหมดเสร็จสิ้น  พร้อมด้วยค่าเสียหายและค่าใช้จ่ายทั้งหลายอันเนื่องจากการผิดนัดชำระหนี้ของผู้กู้ รวมทั้งค่าใช้จ่าย\r\nในการเตือน เรียกร้อง ทวงถาม ดำเนินคดีและการบังคับชำระหนี้จนเต็มจำนวน\r\n"));

                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 6. การเปิดเผยข้อมูล"));
                body.AppendChild(NormalParagraphWith_2Tabs("ในการวิเคราะห์ข้อมูลเพื่อประกอบการพิจารณาให้สินเชื่อ การแก้ไขหนี้ หรือการปรับปรุงโครงสร้างหนี้ของผู้ให้กู้แก่ผู้กู้นั้น ผู้กู้ตกลงยินยอมให้ผู้ให้กู้ตรวจสอบและใช้ข้อมูลเกี่ยวกับการเงิน ประวัติและภาระหนี้  ที่ผู้กู้มีอยู่กับสถาบันการเงิน และนิติบุคคลอื่น รวมทั้งข้อมูลเครดิตของผู้กู้ที่ได้ถูกรวบรวมไว้ที่ บริษัท ข้อมูลเครดิตแห่งชาติ จำกัด  หรือบริษัทข้อมูลเครดิตใด ๆ ตามพระราชบัญญัติการประกอบธุรกิจข้อมูลเครดิต ตลอดจนการตรวจสอบการล้มละลายและหรือ \r\nการบังคับคดีขายทอดตลาดของผู้กู้ได้ โดยไม่ต้องคำนึงว่าผู้กู้จะได้รับอนุมัติสินเชื่อ ไม่ว่าจะเป็นการให้วงเงินสินเชื่อ การแก้ไขหนี้ หรือการปรับปรุงโครงสร้างหนี้จากผู้ให้กู้หรือไม่ก็ตาม\r\n"));

                body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 7. อื่นๆ"));
                body.AppendChild(NormalParagraphWith_2Tabs("7.1 ในระหว่างและตลอดระยะเวลาการกู้เงินตามสัญญานี้ ผู้กู้ยินยอมให้ผู้ให้กู้ หรือตัวแทนผู้ให้กู้เข้าไปตรวจสอบกิจการ ตลอดจนเอกสารหลักฐานทางบัญชีของกิจการ สรรพสมุดและเอกสารอื่นๆ ของผู้กู้ได้ตลอด"));
                body.AppendChild(NormalParagraphWith_2Tabs("7.2 คู่สัญญาตกลงให้ถือเอาเอกสารที่แนบท้ายสัญญานี้ บันทึกข้อตกลง และบรรดาข้อสัญญาต่างๆ เป็นส่วนหนึ่งของสัญญานี้ที่มีผลผูกพันให้ผู้กู้จะต้องปฏิบัติตาม" +
                 " ซึ่งเอกสารแนบท้ายนี้อาจจะทำเพิ่มเติมในภายหลังจากวันทำสัญญานี้ โดยให้ถือเป็นส่วนหนึ่งของสัญญานี้เช่นกัน และหากเอกสารแนบท้ายสัญญาขัดหรือแย้งกันผู้กู้ตกลงปฏิบัติตามคำวินิจฉัยของผู้ให้กู้"));
                body.AppendChild(NormalParagraphWith_2Tabs("7.3 บรรดาหนังสือ จดหมาย คำบอกกล่าวใดๆ เช่น การทวงถาม การบอกเลิกสัญญา ของผู้ให้กู้ที่ส่งไปยังสถานที่ที่ระบุไว้ว่าเป็นที่อยู่ของผู้กู้ข้างต้น" +
                 "หรือสถานที่อยู่ที่ผู้กู้แจ้งเปลี่ยนแปลง โดยส่งเองหรือส่งทางไปรษณีย์ลงทะเบียน หรือไม่ลงทะเบียนไม่ว่าจะมีผู้รับไว้หรือไม่มีผู้ใดยอมรับไว้" +
                 "หรือส่งไม่ได้เพราะผู้กู้ย้ายสถานที่อยู่ไปโดยมิได้แจ้งให้ผู้ให้กู้ทราบให้ไว้นั้นหาไม่พบ หรือถูกรื้อถอนทำลายทุกๆ กรณีดังกล่าวให้ถือว่าผู้กู้ได้รับโดยชอบแล้ว"));
                body.AppendChild(NormalParagraphWith_2Tabs("7.4 การสละสิทธิ์ตามสัญญานี้ ในคราวหนึ่งคราวใดของผู้ให้กู้ หรือการที่ผู้ให้กู้มิได้ใช้สิทธิ์ที่มีอยู่ ไม่ถือเป็นการสละสิทธิ์ของผู้ให้กู้ในคราวต่อไปและไม่มีผลกระทบต่อการใช้สิทธิของผู้ให้กู้ในคราวต่อไป"));
                body.AppendChild(NormalParagraphWith_2Tabs("7.5 หากข้อกำหนด และ/หรือเงื่อนไขข้อใดข้อหนึ่งของสัญญานี้ตกเป็นโมฆะ หรือใช้บังคับไม่ได้ตามกฎหมาย ให้ข้อกำหนดและเงื่อนไขอื่น ๆ ยังคงมีผลใช้บังคับได้ต่อไปได้ โดยแยกต่างหากจากส่วนที่เป็นโมฆะหรือไม่สมบูรณ์นั้น"));
                body.AppendChild(NormalParagraphWith_2Tabs("ผู้กู้ได้ตรวจ อ่าน และเข้าใจข้อความในสัญญานี้โดยละเอียดโดยตลอดแล้ว เห็นว่าถูกต้องตามเจตนาทุกประการ จึงได้ลงลายมือชื่อพร้อมประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยาน ณ วัน เดือน ปี ที่ระบุไว้ข้างต้น"));

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
                body.AppendChild(NormalParagraphWith_2Tabs("ข้าพเจ้า………………………………………………………………………………………………………………………………………….\r\nขอรับรองว่าสถานภาพการสมรสของข้าพเจ้าปัจจุบันมีสถานะ\r\n"));
                body.AppendChild(NormalParagraphWith_2Tabs("“ข้าพเจ้าขอรับรองว่าสถานภาพการสมรสที่แจ้งในหนังสือฉบับนี้เป็นความจริงทุกประการหากไม่เป็นความจริงแล้ว ความเสียหายใด ๆ ที่จะเกิดกับผู้ให้การอุดหนุน ข้าพเจ้ายินยอมรับผิดชดใช้ให้แก่ผู้ให้การอุดหนุนทั้งสิ้น”"));
                body.AppendChild(CenteredParagraph("ลงชื่อ.............................................................รับรอง"));
                body.AppendChild(CenteredParagraph("(............................................................)"));
                body.AppendChild(CenteredParagraph("ลงชื่อ....................................................พยาน          ลงชื่อ ........................................................พยาน"));
                body.AppendChild(CenteredParagraph("(............................................................)                 (.........................................................)"));

                body.AppendChild(EmptyParagraph());
                body.AppendChild(RightParagraph("........................................................./ผู้พิมพ์"));
                body.AppendChild(RightParagraph("........................................................./ผู้ตรวจ"));


                // --- Add header for first page (empty) ---
                AddHeaderWithPageNumber(mainPart, body);
            }
            stream.Position = 0;
            return File(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สสว.สัญญาเงินกู้ยืมโครงการพลิกฟื้นวิสาห.docx");
        }
        #endregion

        #region สัญญาจ้างลูกจ้าง

        public IActionResult OnGetWordContactHireEmployee()
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

                // --- Logo section: large, centered, with whitespace above and below ---
                var imagePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
                if (System.IO.File.Exists(imagePath))
                {
                    // Add empty paragraph above logo for spacing
                    //  body.AppendChild(EmptyParagraph());

                    var imagePart = mainPart.AddImagePart(ImagePartType.Png);
                    using (var imgStream = new FileStream(imagePath, FileMode.Open))
                    {
                        imagePart.FeedData(imgStream);
                    }
                    // Make logo larger (e.g., 240x80 px)
                    var element = CreateImage(mainPart.GetIdOfPart(imagePart), 240, 80);
                    var logoPara = new Paragraph(
                        new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                        element
                    );
                    body.AppendChild(logoPara);
                }
                // 2. Document title and subtitle

                body.AppendChild(CenteredBoldColoredParagraph("สัญญาจ้างลูกจ้าง", "000000", "36"));

                body.AppendChild(NormalParagraphWith_2Tabs("สัญญาฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เลขที่ 21 ถนนวิภาวดีรังสิต เขตจตุจักร กรุงเทพมหานคร เมื่อวันที่ {param1}", null, "32"));
                body.AppendChild(NormalParagraphWith_2Tabs("ระหว่าง สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม โดย........................................." +
                    "\r\nผู้อำนวยการฝ่ายศูนย์ให้บริการ SMEs ครบวงจร สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ผู้รับมอบหมายตามคำสั่งสำนักงานฯ ที่ 629/2564 ลงวันที่ 30 กันยายน 2564 ซึ่งต่อไปในสัญญานี้จะเรียกว่า “ผู้ว่าจ้าง”\r\n", null, "32"));
                body.AppendChild(NormalParagraphWith_2Tabs("ฝ่ายหนึ่ง กับ .................................. เลขประจำตัวประชาชน ........................... อยู่บ้านเลขที่ ......................................... " +
                    "ซึ่งต่อไปในสัญญานี้จะเรียกว่า “ลูกจ้าง” อีกฝ่ายหนึ่ง โดยทั้งสองฝ่ายได้ตกลงทำร่วมกันดังมีรายละเอียดต่อไปนี้", null, "32"));

                body.AppendChild(NormalParagraphWith_2Tabs("1.ผู้ว่าจ้างตกลงจ้างลูกจ้างปฏิบัติงานกับผู้ว่าจ้าง โดยให้ปฏิบัติงานภายใต้งาน....................................  ในตำแหน่ง ................................................................... ปฏิบัติหน้าที่ ณ ศูนย์กลุ่มจังหวัดให้บริการ SME ครบวงจร ......................................... " +
                    "โดยมีรายละเอียดหน้าที่ความรับผิดชอบปรากฏตามเอกสารแนบท้ายสัญญาจ้าง ตั้งแต่วันที่ ................................ ถึงวันที่ .........................................", null, "32"));
                body.AppendChild(NormalParagraphWith_2Tabs("2.ผู้ว่าจ้างจะจ่ายค่าจ้างให้แก่ลูกจ้างในระหว่างระยะเวลาการปฏิบัติงานของลูกจ้างตามสัญญานี้ในอัตราเดือนละ .........................................บาท (.........................................)" +
                    "โดยจะจ่ายให้ในวันทำการก่อนวันทำการสุดท้ายของธนาคารในเดือนนั้นสามวันทำการ และนำเข้าบัญชีเงินฝากของลูกจ้าง ณ ที่ทำการของผู้ว่าจ้าง หรือ ณ ที่อื่นใดตามที่ผู้ว่าจ้างกำหนด", null, "32"));
                body.AppendChild(NormalParagraphWith_2Tabs("3.ในการจ่ายค่าจ้าง และ/หรือ เงินในลักษณะอื่นให้แก่ลูกจ้าง ลูกจ้างตกลงยินยอมให้ผู้ว่าจ้างหักภาษี ณ ที่จ่าย และ/หรือ เงินอื่นใดที่ต้องหักโดยชอบด้วยระเบียบ ข้อบังคับของผู้ว่าจ้างหรือตามกฎหมายที่เกี่ยวข้อง", null, "32"));
                body.AppendChild(NormalParagraphWith_2Tabs("4.ตลอดระยะเวลาการปฏิบัติงานตามสัญญานี้ ลูกจ้างมีสิทธิได้รับสิทธิประโยชน์อื่น ๆ ตามที่กำหนดไว้ใน ระเบียบ ข้อบังคับ คำสั่ง หรือประกาศใด ๆ ตามที่ผู้ว่าจ้างกำหนด", null, "32"));
                body.AppendChild(NormalParagraphWith_2Tabs("5.ผู้ว่าจ้างจะทำการประเมินผลการปฏิบัติงานอย่างน้อยปีละสองครั้ง ตามหลักเกณฑ์และวิธีการที่ผู้ว่าจ้างกำหนด ทั้งนี้ หากผลการประเมินไม่ผ่านตามหลักเกณฑ์ที่กำหนด ผู้ว่าจ้างมีสิทธิบอกเลิกสัญญาจ้างได้ และลูกจ้างไม่มีสิทธิเรียกร้องเงินชดเชยหรือเงินอื่นใด", null, "32"));
                body.AppendChild(NormalParagraphWith_2Tabs("6.ตลอดระยะเวลาการปฏิบัติงานตามสัญญานี้ ลูกจ้างจะต้องปฏิบัติตามกฎ ระเบียบ ข้อบังคับ คำสั่งหรือประกาศใด ๆ ของผู้ว่าจ้าง " +
                    "ตลอดจนมีหน้าที่ต้องรักษาวินัยและยอมรับการลงโทษทางวินัยของผู้ว่าจ้างโดยเคร่งครัด และยินยอมให้ถือว่า กฎหมาย ระเบียบ ข้อบังคับ หรือคำสั่งต่าง ๆ ของผู้ว่าจ้างเป็นส่วนหนึ่งของสัญญาจ้างนี้", null, "32"));

                body.AppendChild(NormalParagraphWith_2Tabs("ในกรณีลูกจ้างจงใจขัดคำสั่งโดยชอบของผู้ว่าจ้างหรือละเลยไม่นำพาต่อคำสั่งเช่นว่านั้นเป็นอาจิณ หรือประการอื่นใด อันไม่สมควรกับการปฏิบัติหน้าที่ของตนให้ลุล่วงไปโดยสุจริตและถูกต้อง ลูกจ้างยินยอมให้ผู้ว่าจ้างบอกเลิกสัญญาจ้างโดยมิต้องบอกกล่าวล่วงหน้า", null, "32"));
                body.AppendChild(NormalParagraphWith_2Tabs("7. ลูกจ้างต้องปฏิบัติงานให้กับผู้ว่าจ้าง ตามที่ได้รับมอบหมายด้วยความซื่อสัตย์ สุจริต และตั้งใจปฏิบัติงานอย่างเต็มกำลังความสามารถของตน โดยแสวงหาความรู้และทักษะเพิ่มเติมหรือกระทำการใด " +
                    "เพื่อให้ผลงานในหน้าที่มีคุณภาพดีขึ้น ทั้งนี้ ต้องรักษาผลประโยชน์และชื่อเสียงของผู้ว่าจ้าง และไม่เปิดเผยความลับหรือข้อมูลของทางราชการให้ผู้หนึ่งผู้ใดทราบ โดยมิได้รับอนุญาตจากผู้รับผิดชอบงานนั้น ๆ", null, "32"));
                body.AppendChild(NormalParagraphWith_2Tabs("8. สัญญานี้สิ้นสุดลงเมื่อเข้ากรณีใดกรณีหนึ่ง ดังต่อไปนี้", null, "32"));

                body.AppendChild(NormalParagraphWith_3Tabs("8.1 สิ้นสุดระยะเวลาตามสัญญาจ้าง", null, "32"));
                body.AppendChild(NormalParagraphWith_3Tabs("8.2 เมื่อผู้ว่าจ้างบอกเลิกสัญญาจ้าง หรือลูกจ้างบอกเลิกสัญญาจ้างตามข้อ 10", null, "32"));
                body.AppendChild(NormalParagraphWith_3Tabs("8.3 ลูกจ้างกระทำการผิดวินัยร้ายแรง", null, "32"));
                body.AppendChild(NormalParagraphWith_3Tabs("8.4 ลูกจ้างไม่ผ่านการประเมินผลการปฏิบัติงานของลูกจ้างตามข้อ 5", null, "32"));
                body.AppendChild(NormalParagraphWith_2Tabs("9. ในกรณีที่สัญญาสิ้นสุดตามข้อ 8.3 และ 8.4 ลูกจ้างยินยอมให้ผู้ว่าจ้างสั่งให้ลูกจ้างพ้นสภาพการเป็นลูกจ้างได้ทันที โดยไม่จำเป็นต้องมีหนังสือว่ากล่าวตักเตือน และผู้ว่าจ้างไม่ต้องจ่ายค่าชดเชยหรือเงินอื่นใดให้แก่ลูกจ้างทั้งสิ้น เว้นแต่ค่าจ้างที่ลูกจ้างจะพึงได้รับตามสิทธิ", null, "32"));
                body.AppendChild(NormalParagraphWith_2Tabs("10. ลูกจ้างมีสิทธิบอกเลิกสัญญาจ้างได้ก่อนสัญญาครบกำหนด โดยทำหนังสือแจ้งเป็นลายลักษณ์อักษรต่อผู้ว่าจ้างได้ทราบล่วงหน้าไม่น้อยกว่า 30 วัน เมื่อผู้ว่าจ้างได้อนุมัติแล้ว ให้ถือว่าสัญญาจ้างนี้ได้สิ้นสุดลง", null, "32"));
                body.AppendChild(NormalParagraphWith_2Tabs("11. ในกรณีที่ลูกจ้างกระทำการใดอันทำให้ผู้ว่าจ้างได้รับความเสียหาย ไม่ว่าเหตุนั้นผู้ว่าจ้างจะนำมาเป็นเหตุบอกเลิกสัญญาจ้างหรือไม่ก็ตาม ผู้ว่าจ้างมีสิทธิจะเรียกร้องค่าเสียหาย และลูกจ้างยินยอมชดใช้ค่าเสียหายตามที่ผู้ว่าจ้างเรียกร้องทุกประการ ", null, "32"));
                body.AppendChild(NormalParagraphWith_2Tabs("12. ลูกจ้างจะต้องไม่เปิดเผยหรือบอกกล่าวอัตราค่าจ้างของลูกจ้างให้แก่บุคคลใดทราบ ไม่ว่าจะ\r\nโดยวิธีใดหรือเวลาใด เว้นแต่จะเป็นการกระทำตามกฎหมายหรือคำสั่งศาล\r\n", null, "32"));
                body.AppendChild(NormalParagraphWith_2Tabs("สัญญาฉบับนี้ได้จัดทำขึ้นเป็นสัญญาอิเล็กทรอนิกส์คู่สัญญาได้อ่านตรวจสอบและทำความเข้าใจข้อความในสัญญาฉบับนี้โดยละเอียดแล้ว จึงได้ลงลายมือชื่ออิเล็กทรอนิกส์ไว้เป็นหลักฐาน ณ วัน เดือน ปี ดังกล่าวข้างต้น และมีพยานรู้ถึงการลงนามของคู่สัญญา และคู่สัญญาต่างฝ่ายต่างเก็บรักษาไฟล์สัญญาอิเล็กทรอนิกส์ฉบับนี้ไว้เป็นหลักฐาน", null, "32"));

                body.AppendChild(EmptyParagraph());
                body.AppendChild(CenteredParagraph("ลงชื่อ....................................................ผู้ว่าจ้าง                    ลงชื่อ ........................................................ลูกจ้าง"));
                body.AppendChild(CenteredParagraph("(............................................................)                                 (.........................................................)"));
                body.AppendChild(CenteredParagraph("ลงชื่อ....................................................ผู้ว่าจ้าง                    ลงชื่อ ........................................................ลูกจ้าง"));
                body.AppendChild(CenteredParagraph("(............................................................)                                 (.........................................................)"));                // next page
                body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

                body.AppendChild(CenteredBoldParagraph("เอกสารแนบท้ายสัญญาจ้างลูกจ้าง", "38"));
                body.AppendChild(CenteredBoldParagraph("งานศูนย์ให้บริการ SMEs ครบวงจร", "38"));


                body.AppendChild(NormalParagraph("หน้าที่ความรับผิดชอบ : เจ้าหน้าที่ศูนย์กลุ่มจังหวัดให้บริการ SMEs ครบวงจร และ", null, "32"));
                body.AppendChild(NormalParagraph("                เจ้าหน้าที่ศูนย์ให้บริการ SMEs ครบวงจร กรุงเทพมหานคร", null, "32"));
                body.AppendChild(NormalParagraphWith_1Tabs("-\tการปรับปรุงข้อมูลผู้ประกอบการ SME (ไม่น้อยกว่า 30 ราย/เดือน)", null, "32"));
                body.AppendChild(NormalParagraphWith_1Tabs("-\tการให้บริการคำปรึกษา แนะนำทางธุรกิจ อาทิเช่น ด้านบัญชี การเงิน การตลาด การบริหารจัดการ การผลิต กฎหมาย เทคโนโลยีสารสนเทศ และอื่น ๆ ที่เกี่ยวข้องทางธุรกิจ (ไม่น้อยกว่า 30 ราย/เดือน)", null, "32"));
                body.AppendChild(NormalParagraphWith_1Tabs("-\tสนับสนุน เสนอแนะแนวทางการแก้ไขปัญหาให้ SME ได้รับประโยชน์ตามมาตรการของภาครัฐ", null, "32"));
                body.AppendChild(NormalParagraphWith_1Tabs("-\tสนับสนุนการพัฒนาเครือข่ายหน่วยงานให้บริการส่งเสริม SME ให้บริการส่งต่อภายใต้หน่วยงานพันธมิตร การติดตามผลและประสานงานแก้ไขปัญหา", null, "32"));
                body.AppendChild(NormalParagraphWith_1Tabs("-\tสนับสนุนนโยบาย มาตรการ และการทำงานของ สสว. ในการสร้าง ประสาน เชื่อมโยงเครือข่ายในพื้นที่ (รูปแบบ Online & Offline) เพื่อสนับสนุนการปฏิบัติงานตามภารกิจ", null, "32"));
                body.AppendChild(NormalParagraphWith_1Tabs("-\tสนับสนุนจัดทำข้อมูล SME จังหวัด เพื่อนำข้อมูลมาใช้ประโยชน์ในการเสนอแนะทางธุรกิจแก่ SME และเชื่อมโยงไปสู่การแก้ปัญหาหรือการจัดทำมาตรการภาครัฐ", null, "32"));
                body.AppendChild(NormalParagraphWith_1Tabs("-\tปฏิบัติงานภายใต้การบังคับบัญชาของผู้จัดการศูนย์กลุ่มจังหวัดฯ หรือ ผู้จัดการศูนย์ให้บริการ SMEs ครบวงจร กรุงเทพมหานคร ตามประกาศ สสว. และเข้าร่วมกิจกรรมต่าง ๆ ", null, "32"));
                body.AppendChild(NormalParagraphWith_1Tabs("-\tกำกับดูแลข้อมูลตาม พ.ร.บ.การคุ้มครองข้อมูลส่วนบุคคล", null, "32"));
                body.AppendChild(NormalParagraphWith_1Tabs("-\tงานอื่น ๆ ตามที่ได้รับมอบหมาย", null, "32"));

                // --- Add footer with page number centered ---
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

                    new FooterReference() { Type = HeaderFooterValues.Default, Id = footerPartId },
                    new PageSize() { Width = 11906, Height = 16838 }, // A4 size
                    new PageMargin() { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720, Gutter = 0 }
                );
                body.AppendChild(sectionProps);
            }
            stream.Position = 0;
            return File(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างลูกจ้าง.docx");
            #endregion
        }
            #region สัญญาจ้างทำของ
    public IActionResult OnGetWordContactToDoThing()
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

        body.AppendChild(CenteredBoldColoredParagraph("แบบสัญญา", "000000", "36"));
        body.AppendChild(CenteredBoldColoredParagraph("สัญญาจ้างทำของ", "000000", "36"));
        // 2. Document title and subtitle
        body.AppendChild(EmptyParagraph());
        body.AppendChild(RightParagraph("สัญญาเลขที่………….…… (1)...........……..……..."));


        body.AppendChild(NormalParagraphWith_2Tabs("สัญญาฉบับนี้ทำขึ้น ณ ………….……..…………………………………………………………………………......."));
        body.AppendChild(JustifiedParagraph("ตำบล/แขวง…………………..………………….………………. อำเภอ/เขต……………………….….……………………………...\r\n" +
          "จังหวัด…….…………………………….………….เมื่อวันที่ ……….……… เดือน …………………….. พ.ศ. ……....……… \r\n" +
          "ระหว่าง……………………………………………………………… (2) ………………………………………………………………………..\r\n" +
          "โดย………...…………….…………………………….……………(3) ………..…………………………………………..…………………ซึ่ง\r\n" +
          "ต่อไปในสัญญานี้เรียกว่า “ผู้ว่าจ้าง” ฝ่ายหนึ่ง กับ…………….…………..…… (4 ก) …………..…………………….ซึ่ง\r\n" +
          "จดทะเบียนเป็นนิติบุคคล ณ ……………………………………………………………………………………….………….……..มี\r\n" +
          "สำนักงานใหญ่อยู่เลขที่ ……………......……ถนน……………….……………..ตำบล/แขวง…….……….…..……….…....\r\n" +
          "อำเภอ/เขต………………….…..…….จังหวัด………..…………………..….โดย………….…………………………………..……...\r\n" +
          "มีอำนาจลงนามผูกพันนิติบุคคลปรากฏตามหนังสือรับรองของสำนักงานทะเบียนหุ้นส่วนบริษัท ……………\r\n" +
          "ลงวันที่………………………………..… (5)(และหนังสือมอบอำนาจลงวันที่ ……………….……..) แนบท้ายสัญญานี้\r\n" +
          "(6)(ในกรณีที่ผู้รับจ้างเป็นบุคคลธรรมดาให้ใช้ข้อความว่า กับ …………………..….… (4 ข) …………………….............\r\n" +
          "อยู่บ้านเลขที่ …………….….…..….ถนน…………………..……..…...……ตำบล/แขวง ……..………………….….…………\r\n" +
          "อำเภอ/เขต…………………….………….…..จังหวัด…………...…..………….……...……. ผู้ถือบัตรประจำตัวประชาชน\r\n" +
          "เลขที่................................ ดังปรากฏตามสำเนาบัตรประจำตัวประชาชนแนบท้ายสัญญานี้) ซึ่งต่อไปใน\r\n" +
          "สัญญานี้เรียกว่า “ผู้รับจ้าง” อีกฝ่ายหนึ่ง", "32"));

        body.AppendChild(NormalParagraphWith_2Tabs("คู่สัญญาได้ตกลงกันมีข้อความดังต่อไปนี้",null,"32"));
        body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 1ข้อตกลงว่าจ้าง", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ผู้ว่าจ้างตกลงจ้างและผู้รับจ้างตกลงรับจ้างทำงาน….………….…… (7) ………...…..……… \r\n" +
          "ณ …..……………................." +
          "ตำบล/แขวง….…………………………….……..อำเภอ/เขต …………..…………..………................\r\n" +
          "จังหวัด……………………….……….….. ตามข้อกำหนดและเงื่อนไขแห่งสัญญานี้รวมทั้งเอกสารแนบท้ายสัญญา", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ผู้รับจ้างตกลงที่จะจัดหาแรงงานและวัสดุ เครื่องมือเครื่องใช้ ตลอดจนอุปกรณ์ต่างๆ ชนิดดีเพื่อใช้ในงานจ้างตามสัญญานี้", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 2\tเอกสารอันเป็นส่วนหนึ่งของสัญญา", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("2.1 ผนวก 1.………..….(รายละเอียดงานจ้าง)…….……..\tจำนวน.…..(…..….….….) หน้า", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("2.2 ผนวก 2……….…....(ใบเสนอราคา)…………….…......\tจำนวน……(………….….) หน้า", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ความใดในเอกสารแนบท้ายสัญญาที่ขัดหรือแย้งกับข้อความในสัญญานี้ ให้ใช้ข้อความ\r\nในสัญญานี้บังคับ และในกรณีที่เอกสารแนบท้ายสัญญาขัดแย้งกันเอง ผู้รับจ้างจะต้องปฏิบัติตามคำวินิจฉัยของผู้ว่าจ้าง คำวินิจฉัยของผู้ว่าจ้างให้ถือเป็นที่สุด และผู้รับจ้างไม่มีสิทธิเรียกร้องค่าจ้าง หรือค่าเสียหาย หรือค่าใช้จ่ายใดๆ เพิ่มเติมจากผู้ว่าจ้างทั้งสิ้น\r\n", null, "32"));

        body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 3\tหลักประกันการปฏิบัติตามสัญญา", null, "32"));
         body.AppendChild(NormalParagraphWith_3Tabs("ในขณะทำสัญญานี้ผู้รับจ้างได้นำหลักประกันเป็น…………….…...…..(8)..………..………" +
           "เป็นจำนวนเงิน……………....บาท(……………..………….) ซึ่งเท่ากับร้อยละ………(9)…..…(…………..………...) ของราคาค่าจ้างตามสัญญา มามอบให้แก่ผู้ว่าจ้างเพื่อเป็นหลักประกันการปฏิบัติตามสัญญานี้", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("(10)กรณีผู้รับจ้างใช้หนังสือค้ำประกันมาเป็นหลักประกันการปฏิบัติตามสัญญา \r\nหนังสือค้ำประกันดังกล่าวจะต้องออกโดยธนาคารที่ประกอบกิจการในประเทศไทย หรือโดยบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจ\r\nค้ำประกันตามประกาศของธนาคารแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบตามแบบที่คณะกรรมการนโยบายการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐกำหนด หรืออาจเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้ และจะต้องมีอายุ\r\nการค้ำประกันตลอดไปจนกว่าผู้รับจ้างพ้นข้อผูกพันตามสัญญานี้\r\n", null, "32"));

        body.AppendChild(NormalParagraphWith_3Tabs("หลักประกันที่ผู้รับจ้างนำมามอบให้ตามวรรคหนึ่ง จะต้องมีอายุครอบคลุมความรับผิด\r\n" +
          "ทั้งปวงของผู้รับจ้างตลอดอายุสัญญา ถ้าหลักประกันที่ผู้รับจ้างนำมามอบให้ดังกล่าวลดลงหรือเสื่อมค่าลง " +
          "หรือมีอายุไม่ครอบคลุมถึงความรับผิดของผู้รับจ้างตลอดอายุสัญญา ไม่ว่าด้วยเหตุใดๆ ก็ตาม รวมถึงกรณี\r\n" +
          "ผู้รับจ้างส่งมอบงานล่าช้าเป็นเหตุให้ระยะเวลาแล้วเสร็จหรือวันครบกำหนดความรับผิดในความชำรุดบกพร่องตามสัญญาเปลี่ยนแปลงไป ไม่ว่าจะเกิดขึ้นคราวใด ผู้รับจ้างต้องหาหลักประกันใหม่หรือหลักประกันเพิ่มเติมให้มีจำนวนครบถ้วนตามวรรคหนึ่งมามอบให้แก่ผู้ว่าจ้างภายใน...............(……………………..….) วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้าง\r\n", null, "32"));

        body.AppendChild(NormalParagraphWith_3Tabs("หลักประกันที่ผู้รับจ้างนำมามอบไว้ตามข้อนี้ ผู้ว่าจ้างจะคืนให้แก่ผู้รับจ้างโดยไม่มีดอกเบี้ยเมื่อผู้รับจ้างพ้นจากข้อผูกพันและความรับผิดทั้งปวงตามสัญญานี้แล้ว", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 4  ค่าจ้างและการจ่ายเงิน ", null, "32")); 
        body.AppendChild(NormalParagraphWith_3Tabs("(11)(ก) สำหรับการจ่ายเงินค่าจ้างให้ผู้รับจ้างเป็นงวด", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ผู้ว่าจ้างตกลงจ่ายและผู้รับจ้างตกลงรับเงินค่าจ้างจำนวนเงิน……………………..บาท(……………………………..…) ซึ่งได้รวมภาษีมูลค่าเพิ่ม จำนวน…………………บาท (......................................) ตลอดจนภาษีอากรอื่นๆ และค่าใช้จ่ายทั้งปวงด้วยแล้ว โดยกำหนดการจ่ายเงินเป็นงวดๆ ดังนี้", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("งวดที่ 1 เป็นจำนวนเงิน………………………...บาท (…………………………………...………….) เมื่อผู้รับจ้างได้ปฏิบัติงาน……………………………………ให้แล้วเสร็จภายใน…………………………………………………..", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("งวดที่ 2 เป็นจำนวนเงิน…….………………...บาท (…………………………………...………….) เมื่อผู้รับจ้างได้ปฏิบัติงาน…………………………..…..……ให้แล้วเสร็จภายใน……………………………………………", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("งวดสุดท้าย เป็นจำนวนเงิน……………..………....บาท (…………………………………...….…..) เมื่อผู้รับจ้างได้ปฏิบัติงานทั้งหมดให้แล้วเสร็จเรียบร้อยตามสัญญาและผู้ว่าจ้างได้ตรวจรับงานจ้างตามข้อ 11 ไว้โดยครบถ้วนแล้ว", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("(12)(ข) สำหรับการจ่ายเงินค่าจ้างให้ผู้รับจ้างครั้งเดียว", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ผู้ว่าจ้างตกลงจ่ายและผู้รับจ้างตกลงรับเงินค่าจ้างจำนวนเงิน……………………..บาท(……………………………..…) ซึ่งได้รวมภาษีมูลค่าเพิ่ม จำนวน…………………บาท (......................................) ตลอดจนภาษีอากรอื่นๆ และค่าใช้จ่ายทั้งปวงด้วยแล้ว เมื่อผู้รับจ้างได้ปฏิบัติงานทั้งหมดให้แล้วเสร็จเรียบร้อยตามสัญญาและผู้ว่าจ้างได้ตรวจรับงานจ้างตามข้อ 11 ไว้โดยครบถ้วนแล้ว ", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("(13)การจ่ายเงินตามเงื่อนไขแห่งสัญญานี้ ผู้ว่าจ้างจะโอนเงินเข้าบัญชีเงินฝากธนาคารของผู้รับจ้าง ชื่อธนาคาร………..………….……….สาขา……………..…….…..ชื่อบัญชี……………….……………………เลขที่บัญชี……………………………ทั้งนี้ ผู้รับจ้างตกลงเป็นผู้รับภาระเงินค่าธรรมเนียมหรือค่าบริการอื่นใดเกี่ยวกับการโอน รวมทั้งค่าใช้จ่ายอื่นใด (ถ้ามี) ที่ธนาคารเรียกเก็บ และยินยอมให้มีการหักเงินดังกล่าว\r\nจากจำนวนเงินโอนในงวดนั้นๆ (ความในวรรคนี้ใช้สำหรับกรณีที่หน่วยงานของรัฐจะจ่ายเงินตรง\r\nให้แก่ผู้รับจ้าง (ระบบ Direct Payment) โดยการโอนเงินเข้าบัญชีเงินฝากธนาคารของผู้รับจ้าง \r\nตามแนวทางที่กระทรวงการคลังหรือหน่วยงานของรัฐเจ้าของงบประมาณเป็นผู้กำหนด แล้วแต่กรณี)\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(14)ข้อ 5 เงินค่าจ้างล่วงหน้า", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ผู้ว่าจ้างตกลงจ่ายเงินค่าจ้างล่วงหน้าให้แก่ผู้รับจ้าง เป็นจำนวนเงิน…………..…..…บาท(………………..….…) ซึ่งเท่ากับร้อยละ……....…(……….…………....) ของราคาค่าจ้างตามสัญญาที่ระบุไว้ในข้อ 4", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("เงินค่าจ้างล่วงหน้าดังกล่าวจะจ่ายให้ภายหลังจากที่ผู้รับจ้างได้วางหลักประกันการรับเงินค่าจ้างล่วงหน้าเป็น...................... (หนังสือค้ำประกันหรือหนังสือค้ำประกันอิเล็กทรอนิกส์ของธนาคาร\r\nภายในประเทศหรือพันธบัตรรัฐบาลไทย) ………………....เต็มตามจำนวนเงินค่าจ้างล่วงหน้านั้นให้แก่ผู้ว่าจ้าง ผู้รับจ้างจะต้องออกใบเสร็จรับเงินค่าจ้างล่วงหน้าตามแบบที่ผู้ว่าจ้างกำหนดให้และผู้รับจ้างตกลงที่จะกระทำตามเงื่อนไขอันเกี่ยวกับการใช้จ่ายและการใช้คืนเงินค่าจ้างล่วงหน้านั้น ดังต่อไปนี้\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("5.1 ผู้รับจ้างจะใช้เงินค่าจ้างล่วงหน้านั้นเพื่อเป็นค่าใช้จ่ายในการปฏิบัติงานตามสัญญาเท่านั้น หากผู้รับจ้างใช้จ่ายเงินค่าจ้างล่วงหน้าหรือส่วนใดส่วนหนึ่งของเงินค่าจ้างล่วงหน้านั้นในทางอื่น ผู้ว่าจ้างอาจจะเรียกเงินค่าจ้างล่วงหน้านั้นคืนจากผู้รับจ้างหรือบังคับเอาจากหลักประกันการรับเงินค่าจ้างล่วงหน้า\r\nได้ทันที\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("5.2 เมื่อผู้ว่าจ้างเรียกร้อง ผู้รับจ้างต้องแสดงหลักฐานการใช้จ่ายเงินค่าจ้างล่วงหน้า เพื่อพิสูจน์ว่าได้เป็นไปตามข้อ 5.1 ภายในกำหนด 15 (สิบห้า) วัน นับถัดจากวันได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้าง \r\nหากผู้รับจ้างไม่อาจแสดงหลักฐานดังกล่าวภายในกำหนด 15 (สิบห้า) วัน ผู้ว่าจ้างอาจเรียกเงินค่าจ้างล่วงหน้าคืนจากผู้รับจ้างหรือบังคับเอาจากหลักประกันการรับเงินค่าจ้างล่วงหน้าได้ทันที\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("5.3 ในการจ่ายเงินค่าจ้างให้แก่ผู้รับจ้างตามข้อ 4 ผู้ว่าจ้างจะหักคืนเงินค่าจ้างล่วงหน้าในแต่ละงวดเพื่อชดใช้คืนเงินค่าจ้างล่วงหน้าไว้จำนวนร้อยละ .............(...........) ของจำนวนเงินค่าจ้างในแต่ละงวดจนกว่าจำนวนเงินที่หักไว้จะครบตามจำนวนเงินที่หักค่าจ้างล่วงหน้าที่ผู้รับจ้างได้รับไปแล้ว ยกเว้นค่าจ้างงวดสุดท้ายจะหักไว้เป็นจำนวนเท่ากับจำนวนเงินค่าจ้างล่วงหน้าที่เหลือทั้งหมด", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("5.4 เงินจำนวนใดๆ ก็ตามที่ผู้รับจ้างจะต้องจ่ายให้แก่ผู้ว่าจ้างเพื่อชำระหนี้หรือ\r\nเพื่อชดใช้ความรับผิดต่างๆ ตามสัญญา ผู้ว่าจ้างจะหักเอาจากเงินค่าจ้างงวดที่จะจ่ายให้แก่ผู้รับจ้าง\r\nก่อนที่จะหักชดใช้คืนเงินค่าจ้างล่วงหน้า\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("5.5 ในกรณีที่มีการบอกเลิกสัญญา หากเงินค่าจ้างล่วงหน้าที่เหลือเกินกว่าจำนวนเงินที่ผู้รับจ้างจะได้รับหลังจากหักชดใช้ในกรณีอื่นแล้ว ผู้รับจ้างจะต้องจ่ายคืนเงินจำนวนที่เหลือนั้นให้แก่\r\nผู้ว่าจ้างภายใน 7 (เจ็ด) วัน นับถัดจากวันได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้าง\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("5.6 ผู้ว่าจ้างจะคืนหลักประกันเงินค่าจ้างล่วงหน้าให้แก่ผู้รับจ้างต่อเมื่อผู้ว่าจ้างได้หักเงินค่าจ้างไว้ครบจำนวนเงินค่าจ้างล่วงหน้าตามข้อ 5.3", null, "32"));
        
        body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 6  กำหนดเวลาแล้วเสร็จและสิทธิของผู้ว่าจ้างในการบอกเลิกสัญญา", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ผู้รับจ้างต้องเริ่มทำงานที่รับจ้างภายในวันที่….... เดือน……………… พ.ศ. ………. และจะต้องทำงานให้แล้วเสร็จบริบูรณ์ภายในวันที่ ….... เดือน …………. พ.ศ. …...…. ถ้าผู้รับจ้างมิได้ลงมือทำงานภายในกำหนดเวลา หรือไม่สามารถทำงานให้แล้วเสร็จตามกำหนดเวลา หรือมีเหตุให้เชื่อได้ว่า\r\nผู้รับจ้างไม่สามารถทำงานให้แล้วเสร็จภายในกำหนดเวลา หรือจะแล้วเสร็จล่าช้าเกินกว่ากำหนดเวลา \r\nหรือผู้รับจ้างทำผิดสัญญาข้อใดข้อหนึ่ง หรือตกเป็นผู้ถูกพิทักษ์ทรัพย์เด็ดขาดหรือตกเป็นผู้ล้มละลาย หรือเพิกเฉยไม่ปฏิบัติตามคำสั่งของคณะกรรมการตรวจรับพัสดุ ผู้ว่าจ้างมีสิทธิที่จะบอกเลิกสัญญานี้ได้ และมีสิทธิจ้างผู้รับจ้างรายใหม่เข้าทำงานของผู้รับจ้างให้ลุล่วงไปได้ด้วย การใช้สิทธิบอกเลิกสัญญานั้นไม่กระทบสิทธิของผู้ว่าจ้างที่จะเรียกร้องค่าเสียหายจากผู้รับจ้าง \r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("การที่ผู้ว่าจ้างไม่ใช้สิทธิเลิกสัญญาดังกล่าวข้างต้นนั้น ไม่เป็นเหตุให้ผู้รับจ้างพ้นจาก\r\nความรับผิดตามสัญญา\r\n", null, "32"));

        body.AppendChild(NormalParagraphWith_2Tabs("(15)ข้อ 7 ความรับผิดชอบในความชำรุดบกพร่องของงานจ้าง", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("เมื่องานแล้วเสร็จบริบูรณ์ และผู้ว่าจ้างได้รับมอบงานจากผู้รับจ้างหรือจากผู้รับจ้างรายใหม่ ในกรณีที่มีการบอกเลิกสัญญาตามข้อ 6 หากมีเหตุชำรุดบกพร่องหรือเสียหายเกิดขึ้นจากการจ้างนี้ \r\nภายในกำหนด.....(16)…..….(……..…..) ปี …….…(……....….) เดือน นับถัดจากวันที่ได้รับมอบงานดังกล่าว \r\nซึ่งความชำรุดบกพร่องหรือเสียหายนั้นเกิดจากความบกพร่องของผู้รับจ้างอันเกิดจากการใช้วัสดุที่ไม่ถูกต้องหรือทำไว้ไม่เรียบร้อย หรือทำไม่ถูกต้องตามมาตรฐานแห่งหลักวิชา ผู้รับจ้างจะต้องรีบทำการแก้ไข\r\nให้เป็นที่เรียบร้อยโดยไม่ชักช้า โดยผู้ว่าจ้างไม่ต้องออกเงินใดๆ ในการนี้ทั้งสิ้น หากผู้รับจ้างไม่กระทำการ\r\nดังกล่าวภายในกำหนด……...(………..……) วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้างหรือไม่ทำการแก้ไขให้ถูกต้องเรียบร้อยภายในเวลาที่ผู้ว่าจ้างกำหนด ให้ผู้ว่าจ้างมีสิทธิที่จะทำการนั้นเอง\r\nหรือจ้างผู้อื่นให้ทำงานนั้น โดยผู้รับจ้างต้องเป็นผู้ออกค่าใช้จ่ายเองทั้งสิ้น\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ในกรณีเร่งด่วนจำเป็นต้องรีบแก้ไขเหตุชำรุดบกพร่องหรือเสียหายโดยเร็ว และไม่อาจรอให้ผู้รับจ้างแก้ไขในระยะเวลาที่กำหนดไว้ตามวรรคหนึ่งได้ ผู้ว่าจ้างมีสิทธิเข้าจัดการแก้ไขเหตุชำรุดบกพร่องหรือเสียหายนั้นเอง หรือจ้างผู้อื่นให้ซ่อมแซมความชำรุดบกพร่องหรือเสียหาย โดยผู้รับจ้าง\r\nต้องรับผิดชอบชำระค่าใช้จ่ายทั้งหมด\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("การที่ผู้ว่าจ้างทำการนั้นเอง หรือจ้างผู้อื่นให้ทำงานนั้นแทนผู้รับจ้าง ไม่ทำให้ผู้รับจ้าง\r\nหลุดพ้นจากความรับผิดตามสัญญา หากผู้รับจ้างไม่ชดใช้ค่าใช้จ่ายหรือค่าเสียหายตามที่ผู้ว่าจ้างเรียกร้องผู้ว่าจ้างมีสิทธิบังคับจากหลักประกันการปฏิบัติตามสัญญาได้\r\n", null, "32"));
       
        body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 8\tการจ้างช่วง", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ผู้รับจ้างจะต้องไม่เอางานทั้งหมดหรือแต่บางส่วนแห่งสัญญานี้ไปจ้างช่วงอีกทอดหนึ่ง เว้นแต่การจ้างช่วงงานแต่บางส่วนที่ได้รับอนุญาตเป็นหนังสือจากผู้ว่าจ้างแล้ว การที่ผู้ว่าจ้างได้อนุญาต\r\nให้จ้างช่วงงานแต่บางส่วนดังกล่าวนั้น ไม่เป็นเหตุให้ผู้รับจ้างหลุดพ้นจากความรับผิดหรือพันธะหน้าที่\r\nตามสัญญานี้ และผู้รับจ้างจะยังคงต้องรับผิดในความผิดและความประมาทเลินเล่อของผู้รับจ้างช่วง \r\nหรือของตัวแทนหรือลูกจ้างของผู้รับจ้างช่วงนั้นทุกประการ\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("กรณีผู้รับจ้างไปจ้างช่วงงานแต่บางส่วนโดยฝ่าฝืนความในวรรคหนึ่ง ผู้รับจ้างต้องชำระค่าปรับให้แก่ผู้ว่าจ้างเป็นจำนวนเงินในอัตราร้อยละ........(17)….....(.........................) ของวงเงินของงาน\r\nที่จ้างช่วงตามสัญญา ทั้งนี้ ไม่ตัดสิทธิผู้ว่าจ้างในการบอกเลิกสัญญา\r\n", null, "32"));
        
        body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 9\tความรับผิดของผู้รับจ้าง", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ผู้รับจ้างจะต้องรับผิดต่ออุบัติเหตุ ความเสียหาย หรือภยันตรายใดๆ อันเกิดจาก\r\nการปฏิบัติงานของผู้รับจ้าง และจะต้องรับผิดต่อความเสียหายจากการกระทำของลูกจ้างหรือตัวแทน\r\nของผู้รับจ้าง และจากการปฏิบัติงานของผู้รับจ้างช่วงด้วย (ถ้ามี)\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ความเสียหายใดๆ อันเกิดแก่งานที่ผู้รับจ้างได้ทำขึ้น แม้จะเกิดขึ้นเพราะเหตุสุดวิสัย\r\nก็ตาม ผู้รับจ้างจะต้องรับผิดชอบโดยซ่อมแซมให้คืนดีหรือเปลี่ยนให้ใหม่โดยค่าใช้จ่ายของผู้รับจ้างเอง เว้นแต่ความเสียหายนั้นเกิดจากความผิดของผู้ว่าจ้าง ทั้งนี้ ความรับผิดของผู้รับจ้างดังกล่าวในข้อนี้จะสิ้นสุดลง\r\nเมื่อผู้ว่าจ้างได้รับมอบงานครั้งสุดท้าย ซึ่งหลังจากนั้นผู้รับจ้างคงต้องรับผิดเพียงในกรณีชำรุดบกพร่อง \r\nหรือความเสียหายดังกล่าวในข้อ 7 เท่านั้น\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ผู้รับจ้างจะต้องรับผิดต่อบุคคลภายนอกในความเสียหายใดๆ อันเกิดจาก\r\nการปฏิบัติงานของผู้รับจ้าง หรือลูกจ้างหรือตัวแทนของผู้รับจ้าง รวมถึงผู้รับจ้างช่วง (ถ้ามี) ตามสัญญานี้ หากผู้ว่าจ้างถูกเรียกร้องหรือฟ้องร้องหรือต้องชดใช้ค่าเสียหายให้แก่บุคคลภายนอกไปแล้ว ผู้รับจ้างจะต้อง\r\nดำเนินการใดๆ เพื่อให้มีการว่าต่างแก้ต่างให้แก่ผู้ว่าจ้างโดยค่าใช้จ่ายของผู้รับจ้างเอง รวมทั้งผู้รับจ้างจะต้องชดใช้ค่าเสียหายนั้นๆ ตลอดจนค่าใช้จ่ายใดๆ อันเกิดจากการถูกเรียกร้องหรือถูกฟ้องร้องให้แก่ผู้ว่าจ้างทันที\r\n", null, "32"));

        body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 10\tการจ่ายเงินแก่ลูกจ้าง", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ผู้รับจ้างจะต้องจ่ายเงินแก่ลูกจ้างที่ผู้รับจ้างได้จ้างมาในอัตราและตามกำหนดเวลา\r\nที่ผู้รับจ้างได้ตกลงหรือทำสัญญาไว้ต่อลูกจ้างดังกล่าว\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ถ้าผู้รับจ้างไม่จ่ายเงินค่าจ้างหรือค่าทดแทนอื่นใดแก่ลูกจ้างดังกล่าวในวรรคหนึ่ง \r\nผู้ว่าจ้างมีสิทธิที่จะเอาเงินค่าจ้างที่จะต้องจ่ายแก่ผู้รับจ้างมาจ่ายให้แก่ลูกจ้างของผู้รับจ้างดังกล่าว และให้ถือว่าผู้ว่าจ้างได้จ่ายเงินจำนวนนั้นเป็นค่าจ้างให้แก่ผู้รับจ้างตามสัญญาแล้ว\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ผู้รับจ้างจะต้องจัดให้มีประกันภัยสำหรับลูกจ้างทุกคนที่จ้างมาทำงาน โดยให้ครอบคลุมถึงความรับผิดทั้งปวงของผู้รับจ้าง รวมทั้งผู้รับจ้างช่วง (ถ้ามี) ในกรณีความเสียหายที่คิดค่าสินไหมทดแทนได้ตามกฎหมาย ซึ่งเกิดจากอุบัติเหตุหรือภยันตรายใดๆ ต่อลูกจ้างหรือบุคคลอื่นที่ผู้รับจ้าง\r\nหรือผู้รับจ้างช่วงจ้างมาทำงาน ผู้รับจ้างจะต้องส่งมอบกรมธรรม์ประกันภัยดังกล่าวพร้อมทั้งหลักฐาน\r\nการชำระเบี้ยประกันให้แก่ผู้ว่าจ้างเมื่อผู้ว่าจ้างเรียกร้อง\r\n", null, "32"));
       
        body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 11\tการตรวจรับงานจ้าง", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("เมื่อผู้ว่าจ้างได้ตรวจรับงานจ้างที่ส่งมอบและเห็นว่าถูกต้องครบถ้วนตามสัญญาแล้ว \r\nผู้ว่าจ้างจะออกหลักฐานการรับมอบเป็นหนังสือไว้ให้ เพื่อผู้รับจ้างนำมาเป็นหลักฐานประกอบการขอรับเงินค่างานจ้างนั้น\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ถ้าผลของการตรวจรับงานจ้างปรากฏว่างานจ้างที่ผู้รับจ้างส่งมอบไม่ตรงตามสัญญา\r\nผู้ว่าจ้างทรงไว้ซึ่งสิทธิที่จะไม่รับงานจ้างนั้น ในกรณีเช่นว่านี้ ผู้รับจ้างต้องทำการแก้ไขให้ถูกต้องตาม\r\nสัญญาด้วยค่าใช้จ่ายของผู้รับจ้างเอง และระยะเวลาที่เสียไปเพราะเหตุดังกล่าวผู้รับจ้างจะนำมาอ้างเป็นเหตุขอขยายเวลาส่งมอบงานจ้างตามสัญญาหรือของดหรือลดค่าปรับไม่ได้\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("(18)ในกรณีที่ผู้รับจ้างส่งมอบงานจ้างถูกต้องแต่ไม่ครบจำนวน หรือส่งมอบครบจำนวน แต่ไม่ถูกต้องทั้งหมด ผู้ว่าจ้างจะตรวจรับงานจ้างเฉพาะส่วนที่ถูกต้อง โดยออกหลักฐานการตรวจรับงานจ้างเฉพาะส่วนนั้นก็ได้ (ความในวรรคสามนี้ จะไม่กำหนดไว้ในกรณีที่ผู้ว่าจ้างต้องการงานจ้างทั้งหมด\r\nในคราวเดียวกัน หรืองานจ้างที่ประกอบเป็นชุดหรือหน่วย ถ้าขาดส่วนประกอบอย่างหนึ่งอย่างใด\r\nไปแล้ว จะไม่สามารถใช้งานได้โดยสมบูรณ์)\r\n", null, "32"));
       
        body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 12\tรายละเอียดของงานจ้างคลาดเคลื่อน", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ผู้รับจ้างรับรองว่าได้ตรวจสอบและทำความเข้าใจในรายละเอียดของงานจ้าง\r\nโดยถี่ถ้วนแล้ว หากปรากฏว่ารายละเอียดของงานจ้างนั้นผิดพลาดหรือคลาดเคลื่อนไปจากหลักการ\r\nทางวิศวกรรมหรือทางเทคนิค ผู้รับจ้างตกลงที่จะปฏิบัติตามคำวินิจฉัยของผู้ว่าจ้าง คณะกรรมการตรวจรับพัสดุ เพื่อให้งานแล้วเสร็จบริบูรณ์ คำวินิจฉัยดังกล่าวให้ถือเป็นที่สุด โดยผู้รับจ้างจะคิดค่าจ้าง ค่าเสียหาย หรือค่าใช้จ่ายใดๆ เพิ่มขึ้นจากผู้ว่าจ้าง หรือขอขยายอายุสัญญาไม่ได้\r\n", null, "32"));
       
        body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 13\tค่าปรับ", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("หากผู้รับจ้างไม่สามารถทำงานให้แล้วเสร็จภายในเวลาที่กำหนดไว้ในสัญญา\r\nและผู้ว่าจ้างยังมิได้บอกเลิกสัญญา ผู้รับจ้างจะต้องชำระค่าปรับให้แก่ผู้ว่าจ้างเป็นจำนวนเงิน\r\nวันละ…..(19)…. บาท (……………...) นับถัดจากวันที่ครบกำหนดเวลาแล้วเสร็จของงานตามสัญญาหรือวันที่\r\nผู้ว่าจ้างได้ขยายเวลาทำงานให้ จนถึงวันที่ทำงานแล้วเสร็จจริง นอกจากนี้ ผู้รับจ้างยอมให้ผู้ว่าจ้าง\r\nเรียกค่าเสียหายอันเกิดขึ้นจากการที่ผู้รับจ้างทำงานล่าช้าเฉพาะส่วนที่เกินกว่าจำนวนค่าปรับดังกล่าวได้อีกด้วย\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ในระหว่างที่ผู้ว่าจ้างยังมิได้บอกเลิกสัญญานั้น หากผู้ว่าจ้างเห็นว่าผู้รับจ้าง\r\nจะไม่สามารถปฏิบัติตามสัญญาต่อไปได้ ผู้ว่าจ้างจะใช้สิทธิบอกเลิกสัญญาและใช้สิทธิตามข้อ 14 ก็ได้ \r\nและถ้าผู้ว่าจ้างได้แจ้งข้อเรียกร้องไปยังผู้รับจ้างเมื่อครบกำหนดเวลาแล้วเสร็จของงานขอให้ชำระค่าปรับแล้ว ผู้ว่าจ้างมีสิทธิที่จะปรับผู้รับจ้างจนถึงวันบอกเลิกสัญญาได้อีกด้วย\r\n", null, "32"));
        
        body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 14\tสิทธิของผู้ว่าจ้างภายหลังบอกเลิกสัญญา", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ในกรณีที่ผู้ว่าจ้างบอกเลิกสัญญา ผู้ว่าจ้างอาจทำงานนั้นเองหรือว่าจ้างผู้อื่นให้ทำงานนั้นต่อจนแล้วเสร็จก็ได้ และในกรณีดังกล่าว ผู้ว่าจ้างมีสิทธิริบหรือบังคับจากหลักประกันการปฏิบัติตามสัญญาทั้งหมดหรือบางส่วนตามแต่จะเห็นสมควร นอกจากนั้น ผู้รับจ้างจะต้องรับผิดชอบในค่าเสียหายซึ่งเป็นจำนวนเกินกว่าหลักประกันการปฏิบัติตามสัญญา รวมทั้งค่าใช้จ่ายที่เพิ่มขึ้นในการทำงานนั้นต่อให้แล้วเสร็จตามสัญญา ซึ่งผู้ว่าจ้างจะหักเอาจากจำนวนเงินใดๆ ที่จะจ่ายให้แก่ผู้รับจ้างก็ได้", null, "32"));
       
        body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 15\tการบังคับค่าปรับ ค่าเสียหาย และค่าใช้จ่าย", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ในกรณีที่ผู้รับจ้างไม่ปฏิบัติตามสัญญาข้อใดข้อหนึ่งด้วยเหตุใดๆ ก็ตาม จนเป็นเหตุ\r\nให้เกิดค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายแก่ผู้ว่าจ้าง ผู้รับจ้างต้องชดใช้ค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายดังกล่าวให้แก่ผู้ว่าจ้างโดยสิ้นเชิงภายในกำหนด.................(....................) วัน นับถัดจากวันที่ได้รับแจ้ง\r\nเป็นหนังสือจากผู้ว่าจ้าง หากผู้รับจ้างไม่ชดใช้ให้ถูกต้องครบถ้วนภายในระยะเวลาดังกล่าวให้ผู้ว่าจ้างมีสิทธิที่จะหักเอาจากจำนวนเงินค่าจ้างที่ต้องชำระ หรือบังคับจากหลักประกันการปฏิบัติตามสัญญาได้ทันที\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("หากค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายที่บังคับจากเงินค่าจ้างที่ต้องชำระ \r\nหรือหลักประกันการปฏิบัติตามสัญญาแล้วยังไม่เพียงพอ ผู้รับจ้างยินยอมชำระส่วนที่เหลือที่ยังขาดอยู่\r\nจนครบถ้วนตามจำนวนค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายนั้น ภายในกำหนด ..................(......................) วัน นับถัดจากวันที่ได้รับแจ้งเป็นหนังสือจากผู้ว่าจ้าง\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("หากมีเงินค่าจ้างตามสัญญาที่หักไว้จ่ายเป็นค่าปรับ ค่าเสียหาย หรือค่าใช้จ่ายแล้ว\r\nยังเหลืออยู่อีกเท่าใด ผู้ว่าจ้างจะคืนให้แก่ผู้รับจ้างทั้งหมด\r\n", null, "32"));
        
        body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 16\tการงดหรือลดค่าปรับ หรือการขยายเวลาปฏิบัติงานตามสัญญา", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ในกรณีที่มีเหตุเกิดจากความผิดหรือความบกพร่องของฝ่ายผู้ว่าจ้าง หรือเหตุสุดวิสัย \r\nหรือเกิดจากพฤติการณ์อันหนึ่งอันใดที่ผู้รับจ้างไม่ต้องรับผิดตามกฎหมาย หรือเหตุอื่นตามที่กำหนด\r\nในกฎกระทรวง ซึ่งออกตามความในกฎหมายว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ ทำให้ผู้รับจ้างไม่สามารถทำงานให้แล้วเสร็จตามเงื่อนไขและกำหนดเวลาแห่งสัญญานี้ได้ ผู้รับจ้างจะต้องแจ้งเหตุหรือพฤติการณ์ดังกล่าวพร้อมหลักฐานเป็นหนังสือให้ผู้ว่าจ้างทราบ เพื่อของดหรือลดค่าปรับ หรือขยายเวลาทำงานออกไปภายใน 15 (สิบห้า) วันนับถัดจากวันที่เหตุนั้นสิ้นสุดลง หรือตามที่กำหนดในกฎกระทรวงดังกล่าว แล้วแต่กรณี\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ถ้าผู้รับจ้างไม่ปฏิบัติให้เป็นไปตามความในวรรคหนึ่ง ให้ถือว่าผู้รับจ้างได้สละสิทธิเรียกร้องในการที่จะของดหรือลดค่าปรับ หรือขยายเวลาทำงานออกไปโดยไม่มีเงื่อนไขใดๆ ทั้งสิ้น เว้นแต่\r\nกรณีเหตุเกิดจากความผิดหรือความบกพร่องของฝ่ายผู้ว่าจ้าง ซึ่งมีหลักฐานชัดแจ้ง หรือผู้ว่าจ้างทราบ\r\nดีอยู่แล้วตั้งแต่ต้น\r\n", null, "32"));

        body.AppendChild(NormalParagraphWith_2Tabs("ข้อ 17\tการใช้เรือไทย", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ในการปฏิบัติตามสัญญานี้ หากผู้รับจ้างจะต้องสั่งหรือนำของเข้ามาจากต่างประเทศรวมทั้งเครื่องมือและอุปกรณ์ที่ต้องนำเข้ามาเพื่อปฏิบัติงานตามสัญญา ไม่ว่าผู้รับจ้างจะเป็นผู้ที่นำของเข้ามาเอง หรือนำเข้ามาโดยผ่านตัวแทนหรือบุคคลอื่นใด ถ้าสิ่งของนั้นต้องนำเข้ามาโดยทางเรือในเส้นทางเดินเรือที่มีเรือไทยเดินอยู่และสามารถให้บริการรับขนได้ตามที่รัฐมนตรีว่าการกระทรวงคมนาคมประกาศกำหนด \r\nผู้รับจ้างต้องจัดการให้สิ่งของดังกล่าวบรรทุกโดยเรือไทยหรือเรือที่มีสิทธิเช่นเดียวกับเรือไทยจากต่างประเทศมายังประเทศไทย เว้นแต่จะได้รับอนุญาตจากกรมเจ้าท่าก่อนบรรทุกของนั้นลงเรืออื่นที่มิใช่เรือไทย\r\nหรือเป็นของที่รัฐมนตรีว่าการกระทรวงคมนาคมประกาศยกเว้นให้บรรทุกโดยเรืออื่นได้ ทั้งนี้ไม่ว่าการสั่งหรือนำเข้าสิ่งของดังกล่าวจากต่างประเทศจะเป็นแบบใด\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ในการส่งมอบงานตามสัญญาให้แก่ผู้ว่าจ้าง ถ้างานนั้นมีสิ่งของตามวรรคหนึ่ง \r\nผู้รับจ้างจะต้องส่งมอบใบตราส่ง (Bill of Lading) หรือสำเนาใบตราส่งสำหรับของนั้น ซึ่งแสดงว่าได้บรรทุกมาโดยเรือไทยหรือเรือที่มีสิทธิเช่นเดียวกับเรือไทยให้แก่ผู้ว่าจ้างพร้อมกับการส่งมอบงานด้วย\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ในกรณีที่สิ่งของดังกล่าวไม่ได้บรรทุกจากต่างประเทศมายังประเทศไทยโดยเรือไทยหรือเรือที่มีสิทธิเช่นเดียวกับเรือไทย ผู้รับจ้างต้องส่งมอบหลักฐานซึ่งแสดงว่าได้รับอนุญาตจากกรมเจ้าท่า ให้บรรทุกของโดยเรืออื่นได้ หรือหลักฐานซึ่งแสดงว่าได้ชำระค่าธรรมเนียมพิเศษเนื่องจากการไม่บรรทุกของโดยเรือไทยตามกฎหมายว่าด้วยการส่งเสริมการพาณิชยนาวีแล้วอย่างใดอย่างหนึ่งแก่ผู้ว่าจ้างด้วย", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ในกรณีที่ผู้รับจ้างไม่ส่งมอบหลักฐานอย่างใดอย่างหนึ่งดังกล่าวในวรรคสอง\r\nและวรรคสามให้แก่ผู้ว่าจ้าง แต่จะขอส่งมอบงานดังกล่าวให้ผู้ว่าจ้างก่อนโดยไม่รับชำระเงินค่าจ้าง ผู้ว่าจ้าง มีสิทธิรับงานดังกล่าวไว้ก่อน และชำระเงินค่าจ้างเมื่อผู้รับจ้างได้ปฏิบัติถูกต้องครบถ้วนดังกล่าวแล้วได้\r\n", null, "32"));
        
        body.AppendChild(NormalParagraphWith_2Tabs("สัญญานี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน คู่สัญญาได้อ่านและเข้าใจ\r\nข้อความ โดยละเอียดตลอดแล้ว จึงได้ลงลายมือชื่อ พร้อมทั้งประทับตรา (ถ้ามี) ไว้เป็นสำคัญต่อหน้าพยาน \r\nและคู่สัญญาต่างยึดถือไว้ฝ่ายละหนึ่งฉบับ\r\n", null, "32"));

        body.AppendChild(EmptyParagraph());


        body.AppendChild(CenteredParagraph("ลงชื่อ........................................................................ผู้ว่าจ้าง"));
        body.AppendChild(CenteredParagraph("(................................................................................)"));
        body.AppendChild(CenteredParagraph("ลงชื่อ........................................................................ผู้ว่าจ้าง"));
        body.AppendChild(CenteredParagraph("(................................................................................)"));     
        body.AppendChild(CenteredParagraph("ลงชื่อ......................................................................พยาน"));
        body.AppendChild(CenteredParagraph("(...............................................................................)"));
        body.AppendChild(CenteredParagraph("ลงชื่อ......................................................................พยาน"));
        body.AppendChild(CenteredParagraph("(...............................................................................)"));

        // next page
        body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

        body.AppendChild(CenteredParagraph("วิธีปฏิบัติเกี่ยวกับสัญญาจ้าง"));
        body.AppendChild(NormalParagraphWith_2Tabs("(1) ให้ระบุเลขที่สัญญาในปีงบประมาณหนึ่งๆ ตามลำดับ", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(2) ให้ระบุชื่อของหน่วยงานของรัฐที่เป็นนิติบุคคล เช่น กรม ก. หรือรัฐวิสาหกิจ ข. เป็นต้น", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(3) ให้ระบุชื่อและตำแหน่งของหัวหน้าหน่วยงานของรัฐที่เป็นนิติบุคคลนั้น หรือผู้ที่ได้รับมอบอำนาจ เช่น นาย ก. อธิบดีกรม……………… หรือ นาย ข. ผู้ได้รับมอบอำนาจจากอธิบดีกรม………………..", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(4) ให้ระบุชื่อผู้รับจ้าง", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ก. กรณีนิติบุคคล เช่น ห้างหุ้นส่วนสามัญจดทะเบียน ห้างหุ้นส่วนจำกัด บริษัทจำกัด", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("ข. กรณีบุคคลธรรมดา ให้ระบุชื่อและที่อยู่", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(5) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(6) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(7) ให้ระบุงานที่ต้องการจ้าง", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(8) “หลักประกัน” หมายถึง หลักประกันที่ผู้รับจ้างนำมามอบไว้แก่หน่วยงานของรัฐ \r\nเมื่อลงนามในสัญญา เพื่อเป็นการประกันความเสียหายที่อาจจะเกิดขึ้นจากการปฏิบัติตามสัญญา \r\nดังนี้\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("(๑)\tเงินสด ", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("(๒)\tเช็คหรือดราฟท์ ที่ธนาคารเซ็นสั่งจ่าย ซึ่งเป็นเช็คหรือดราฟท์ลงวันที่ที่ใช้เช็ค\r\nหรือดราฟท์นั้นชำระต่อเจ้าหน้าที่ หรือก่อนวันนั้นไม่เกิน ๓ วันทำการ \r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("(๓)\tหนังสือคํ้าประกันของธนาคารภายในประเทศตามตัวอย่างที่คณะกรรมการนโยบายกําหนด โดยอาจกำหนดเป็นหนังสือค้ำประกันอิเล็กทรอนิกส์ตามวิธีการที่กรมบัญชีกลางกำหนดก็ได้", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("(๔)\tหนังสือค้ำประกันของบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์ที่ได้รับอนุญาตให้ประกอบกิจการเงินทุนเพื่อการพาณิชย์และประกอบธุรกิจค้ำประกันตามประกาศของธนาคาร\r\nแห่งประเทศไทย ตามรายชื่อบริษัทเงินทุนที่ธนาคารแห่งประเทศไทยแจ้งเวียนให้ทราบ โดยอนุโลม\r\nให้ใช้ตามตัวอย่างหนังสือค้ำประกันของธนาคารที่คณะกรรมการนโยบายกําหนด\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_3Tabs("(๕)\tพันธบัตรรัฐบาลไทย", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(9) ให้กำหนดจำนวนเงินหลักประกันการปฏิบัติตามสัญญาตามระเบียบกระทรวงการคลัง\r\nว่าด้วยหลักเกณฑ์การจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. 2560 ข้อ 168\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(10) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(11) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(12) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(13) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(14) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(15) เป็นข้อความหรือเงื่อนไขเพิ่มเติม ซึ่งหน่วยงานของรัฐผู้ทำสัญญาอาจเลือกใช้หรือตัดออกได้ตามข้อเท็จจริง", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(16) กำหนดเวลาที่ผู้รับจ้างจะรับผิดในความชำรุดบกพร่อง โดยปกติจะต้องกำหนด\r\nไม่น้อยกว่า 1 ปี นับถัดจากวันที่ผู้รับจ้างได้รับมอบงานจ้าง หรือกำหนดตามความเหมาะสม\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(17) อัตราค่าปรับตามสัญญาข้อ 8 กรณีผู้รับจ้างไปจ้างช่วงบางส่วนโดยไม่ได้รับอนุญาต\r\nจากผู้ว่าจ้าง ต้องกำหนดค่าปรับเป็นจำนวนเงินไม่น้อยกว่าร้อยละสิบของวงเงินของงานที่จ้างช่วงตามสัญญา\r\n", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(18) ความในวรรคนี้ จะไม่กำหนดไว้ในกรณีที่ผู้ว่าจ้างต้องการสิ่งของทั้งหมดในคราวเดียวกัน หรืองานจ้างที่ประกอบเป็นชุดหรือหน่วย ถ้าขาดส่วนประกอบอย่างหนึ่งอย่างใดไปแล้ว จะไม่สามารถใช้งานได้โดยสมบูรณ์", null, "32"));
        body.AppendChild(NormalParagraphWith_2Tabs("(19) อัตราค่าปรับตามสัญญาข้อ 13 ให้กำหนด ตามระเบียบกระทรวงการคลังว่าด้วยหลักเกณฑ์การจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. 2560 ข้อ 162 ส่วนกรณีจะปรับร้อยละเท่าใด ให้อยู่ในดุลพินิจของหน่วยงานของรัฐผู้ว่าจ้างที่จะพิจารณาแต่ทั้งนี้การที่จะกำหนดค่าปรับเป็นร้อยละเท่าใด จะต้องกำหนดไว้ในเอกสารเชิญชวนด้วย", null, "32"));
       
        AddHeaderWithPageNumber(mainPart, body);
        
      }
      stream.Position = 0;
      return File(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างทำของ.docx");
    }
        #endregion
    }
}
