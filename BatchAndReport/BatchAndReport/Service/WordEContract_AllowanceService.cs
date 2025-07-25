using BatchAndReport.Entities;
using BatchAndReport.Models;
using BatchAndReport.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

public class WordEContract_AllowanceService : IWordEContract_AllowanceService
{
    public byte[] ConvertWordToPdf(byte[] wordBytes)
    {
        try
        {
            using var inputStream = new MemoryStream(wordBytes);
            var doc = new Spire.Doc.Document(); // ✅ ใช้ชื่อเต็มป้องกันชนกับ OpenXML.Document
            doc.LoadFromStream(inputStream, Spire.Doc.FileFormat.Docx);

            using var outputStream = new MemoryStream();
            doc.SaveToStream(outputStream, Spire.Doc.FileFormat.PDF);
            return outputStream.ToArray();
        }
        catch (Exception ex)
        {
            throw new ApplicationException("ConvertWordToPdf failed: " + ex.Message, ex);
        }
    }

    #region 
    // This is your specific handler for the contract report
    public byte[] GenerateWordContactAllowance()
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
            body.AppendChild(CenteredBoldColoredParagraph("สัญญารับเงินอุดหนุน", "0000FF")); // Blue
            body.AppendChild(CenteredBoldColoredParagraph("ตามแนวทางการดำเนินโครงการวิสาหกิจขนาดกลางและขนาดย่อมต่อเนื่อง", "FF0000")); // Red
            body.AppendChild(CenteredParagraph("ที่ศูนย์ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม"));
            body.AppendChild(CenteredParagraph("วันที่...................................................."));

            // 3. Fillable lines (using underlines)
            body.AppendChild(EmptyParagraph());
            body.AppendChild(NormalParagraph("ข้าพเจ้า ...................................................................................................................."));
            body.AppendChild(NormalParagraph("อายุ ......... ปี สัญชาติ .................. สำนักงาน/บ้านตั้งอยู่เลขที่.................. อาคาร..........................................."));
            body.AppendChild(NormalParagraph("หมู่ที่...........ตรอก/ซอย..........................ถนน...........................ตำบล/แขวง.................. ..................."));
            body.AppendChild(NormalParagraph("เขต/อำเภอ................... จังหวัด.................ทะเบียนนิติบุคคลเลขที่/เลขประจำตัวประชาชนที่............................................"));
            body.AppendChild(NormalParagraph("จดทะเบียนเป็นนิติบุคคลเมื่อวันที่ .........................................."));

            // 4. Main body (sample)
            body.AppendChild(EmptyParagraph());
            body.AppendChild(NormalParagraph("ซึ่งต่อไปนี้จะเรียกบุคคลผู้มีนามตามที่ระบุข้างต้นทั้งหมดว่า \"ผู้รับการอุดหนุน\" ได้ทำสัญญาฉบับนี้ให้ไว้แก่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งต่อไปนี้จะเรียกว่า \"ผู้ให้การอุดหนุน\" โดยมีสาระสำคัญดังนี้"));
            body.AppendChild(NormalParagraph("ข้อ 1.  ผู้รับการอุดหนุนได้ขอรับความช่วยเหลือผ่านการอุดหนุนตามมาตรการฟื้นฟูกิจการวิสาหกิจ\r\nขนาดกลางและขนาดย่อมจากผู้ให้การอุดหนุนเป็นจำนวนเงิน ..................... บาท (...........................) ปลอดการชำระเงินต้น ................. เดือน โดยไม่มีดอกเบี้ย แต่มีภาระต้องชำระคืนเงินต้น \r\n"));
            body.AppendChild(NormalParagraph("ข้อ 2. ผู้ให้การอุดหนุนจะให้ความช่วยเหลือด้วยการให้เงินอุดหนุนแก่ผู้รับการอุดหนุน ด้วยการนำเงินหรือโอนเงินเข้าบัญชีธนาคารกรุงไทย จำกัด (มหาชน) สาขา ..................................... เลขที่บัญชี ................................................. ชื่อบัญชี .......................... ซึ่งเป็นบัญชีของผู้รับการอุดหนุน จำนวนเงิน ........................... บาท (..............................) และให้ถือว่าผู้รับการอุดหนุนได้รับเงินอุดหนุนตามสัญญานี้ไปจากผู้ให้การอุดหนุนแล้ว ในวันที่เงินเข้าบัญชีของผู้รับการอุดหนุนดังกล่าว"));
            body.AppendChild(NormalParagraph("ข้อ 3. ห้ามผู้รับการอุดหนุนนำเงินอุดหนุนไปชำระหนี้เดิมที่มีอยู่ก่อนทำสัญญานี้"));
            body.AppendChild(NormalParagraph("ข้อ 4. ผู้รับการอุดหนุนยินยอมให้ธนาคารกรุงไทย จำกัด (มหาชน) ซึ่งกระทำการแทนผู้ให้การอุดหนุน\r\nหักเงินอุดหนุนที่จะได้จากผู้ให้การอุดหนุนเป็นค่าใช้จ่ายหรือค่าธรรมเนียมในการโอนเงินเข้าบัญชีของผู้รับการอุดหนุน\r\nซึ่งธนาคารกรุงไทย จำกัด (มหาชน) เรียกเก็บตามระเบียบของธนาคารได้ โดยไม่ต้องบอกกล่าวหรือแจ้งให้ผู้รับการอุดหนุนทราบล่วงหน้า และให้ถือว่าผู้รับการอุดหนุนได้รับเงินตามจำนวนที่เบิกไปครบถ้วนแล้ว\r\n"));
            body.AppendChild(NormalParagraph("ข้อ 5. ผู้รับการอุดหนุนตกลงผ่อนชำระเงินต้นคืนให้แก่ผู้ให้การอุดหนุนเป็นรายเดือน (งวด) ๆ ละ ไม่น้อยกว่า ....................... บาท (.....................................) ด้วยการโอนเข้าบัญชีตามที่ระบุไว้ในข้อ 2 โดยชำระเงินต้นงวดแรกในเดือนที่ ....................... นับถัดจากวันที่ได้รับเงินอุดหนุน และงวดถัดไปทุกวันที่ .................. ของเดือนจนกว่าจะชำระเสร็จสิ้น \r\nแต่ทั้งนี้จะต้องชำระให้เสร็จสิ้นไม่เกินกว่า .............. ปี (...........) นับแต่วันที่ได้รับเงินอุดหนุน\r\n"));
            body.AppendChild(NormalParagraph("ข้อ 6 การชำระเงินคืนตามข้อ 5 ผู้รับการอุดหนุนตกลงจะนำเงินเข้าบัญชีเงินฝากของผู้รับการอุดหนุน\r\nที่เปิดบัญชีไว้กับธนาคารกรุงไทย จำกัด (มหาชน) ตามข้อ 2 โดยผู้รับการอุดหนุนยินยอมให้ ธนาคารกรุงไทย จำกัด (มหาชน) ซึ่งดำเนินการแทนผู้ให้การอุดหนุน หักเงินจากบัญชีของผู้รับการอุดหนุนดังกล่าวเพื่อชำระคืนเงินอุดหนุนแก่ผู้ให้การอุดหนุน\r\n"));







            body.AppendChild(NormalParagraph(""));
            // 5. Section properties (A4, margins)
            var sectionProps = new SectionProperties(
                new PageSize() { Width = 11906, Height = 16838 },
                new PageMargin() { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 }
            );
            body.AppendChild(sectionProps);

            // Ensure the document is saved before closing
            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream.ToArray();
    }
    // Helper for colored, bold, centered paragraph
    private static Paragraph CenteredBoldColoredParagraph(string text, string hexColor) =>
        new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
            new Run(
                new RunProperties(
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "32" },
                    new Bold(),
                    new DocumentFormat.OpenXml.Office2013.Word.Color { Val = hexColor }
                ),
                new Text(text)
            )
        );

    #endregion

    void AddLogo(MainDocumentPart mainPart, Body body, string imagePath)
    {
        var imagePart = mainPart.AddImagePart(ImagePartType.Png);
        using (var stream = new FileStream(imagePath, FileMode.Open))
        {
            imagePart.FeedData(stream);
        }

        string imagePartId = mainPart.GetIdOfPart(imagePart);

        var drawing = new Drawing(
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = 990000L, Cy = 792000L },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent()
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties()
                {
                    Id = (UInt32Value)1U,
                    Name = "SME Logo"
                },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                    new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true }),
                new DocumentFormat.OpenXml.Drawing.Graphic(
                    new DocumentFormat.OpenXml.Drawing.GraphicData(
                        new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties()
                                {
                                    Id = (UInt32Value)0U,
                                    Name = "logo_SME.png"
                                },
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                            new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                new DocumentFormat.OpenXml.Drawing.Blip()
                                {
                                    Embed = imagePartId,
                                    CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print
                                },
                                new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                            new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                new DocumentFormat.OpenXml.Drawing.Transform2D(
                                    new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                    new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 990000L, Cy = 792000L }),
                                new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                    new DocumentFormat.OpenXml.Drawing.AdjustValueList())
                                { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }))
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            }
        );

        var paragraph = new Paragraph(
            new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center },
                new SpacingBetweenLines() { After = "200" }),
            new Run(drawing));

        body.Append(paragraph);
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
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "32" } // Correct namespace and usage
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
                 new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "32" } // Correct namespace and usage
             ),
             new Text(text)
         )
     );

    // Fix for CS0117: 'FontSize' does not contain a definition for 'Val'
    // The issue arises because the incorrect namespace or type is being used for FontSize.
    // Replace the problematic line with the correct usage of FontSize from DocumentFormat.OpenXml.Wordprocessing.

    private static Paragraph NormalParagraph(string text, JustificationValues? align = null) =>
        new Paragraph(
            align != null ? new ParagraphProperties(new Justification { Val = align.Value }) : null,
            new Run(
                new RunProperties(
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "32" } // Correct namespace and usage
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

    private TableCell CreateSignatureCell(string text)
    {
        return new TableCell(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Right },
                new SpacingBetweenLines { After = "200" }),
            new Run(
                new RunProperties(
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                    new FontSize { Val = "32" }),
                new Text(text))));
    }

}
