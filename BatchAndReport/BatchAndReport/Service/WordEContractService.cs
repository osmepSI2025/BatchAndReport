using BatchAndReport.Entities;
using BatchAndReport.Models;
using BatchAndReport.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.IO;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

public class WordEContractService : IWordEContractService
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
    public byte[] GenJointContractAgreement(ConJointContractModels model)
    {
        using var stream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());
            //AddLogo(mainPart, body, "wwwroot/images/logo_SME.png"); // หรือ path ที่คุณเก็บโลโก้

            ParagraphProperties Centered = new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "200" },
                new Indentation { Left = "0", Right = "0" });

            ParagraphProperties NormalLeft = new ParagraphProperties(
                new Justification { Val = JustificationValues.Left },
                new SpacingBetweenLines { After = "200" },
                new Indentation { Left = "400", Hanging = "0" });

            var defaultFont = new RunProperties(
                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                new FontSize { Val = "32" });

            Paragraph CreateParagraph(string text, ParagraphProperties props = null, bool bold = false)
            {
                var runProps = (RunProperties)defaultFont.CloneNode(true);
                if (bold) runProps.AppendChild(new Bold());
                return new Paragraph(
                    props != null ? (ParagraphProperties)props.CloneNode(true) : (ParagraphProperties)NormalLeft.CloneNode(true),
                    new Run(runProps, new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
            }

            body.Append(CreateParagraph("สัญญาร่วมดำเนินการ", Centered, true));
            body.Append(CreateParagraph($"โครงการ {model.ProjectName}", Centered));
            body.Append(CreateParagraph("ระหว่าง", Centered));
            body.Append(CreateParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม", Centered, true));
            body.Append(CreateParagraph($"กับ {model.AgencyName}", Centered));

            body.Append(CreateParagraph($"สัญญาร่วมดำเนินการฉบับนี้ทำขึ้น ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เมื่อวันที่ {model.SignDay} เดือน {model.SignMonth} พ.ศ. {model.SignYear} ระหว่าง"));
            body.Append(CreateParagraph($"สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม โดย {model.SMEOfficialName} ตำแหน่ง {model.SMEOfficialPosition} ซึ่งต่อไปเรียกว่า \"สสว.\" ฝ่ายหนึ่ง กับ \"{model.AgencyName}\" โดย {model.AgencyRepresentative} ตำแหน่ง {model.AgencyPosition} ซึ่งต่อไปในสัญญานี้จะเรียกว่า \"ชื่อหน่วยร่วม\" อีกฝ่ายหนึ่ง"));

            body.Append(CreateParagraph("วัตถุประสงค์ตามสัญญาร่วมดำเนินการ", Centered, true));
            body.Append(CreateParagraph($"คู่สัญญาทั้งสองฝ่ายมีความประสงค์ที่จะร่วมมือกันเพื่อดำเนินการภายใต้โครงการ {model.ProjectName} ซึ่งต่อไปในสัญญานี้จะเรียกว่า \"โครงการ\" โดยมีรายละเอียดโครงการ แผนการดำเนินงาน แผนการใช้จ่ายเงิน และบรรดาเอกสารแนบท้ายสัญญาฉบับนี้ ซึ่งให้ถือเป็นส่วนหนึ่งของสัญญาฉบับนี้ โดยมีวัตถุประสงค์ในการดำเนินโครงการ ดังนี้"));

            foreach (var obj in model.Objectives)
                body.Append(CreateParagraph($"{obj.Number} {obj.Description}"));

            body.Append(CreateParagraph("ข้อ 1 ขอบเขตหน้าที่ของ “สสว.”", null, true));
            foreach (var paragraph in model.SMEDuties)
                body.Append(CreateParagraph(paragraph));

            body.Append(CreateParagraph("ข้อ 2 ขอบเขตหน้าที่ของ “ชื่อหน่วยร่วม”", null, true));
            foreach (var paragraph in model.AgencyDuties)
                body.Append(CreateParagraph(paragraph));

            body.Append(CreateParagraph("ข้อ 3 อื่น ๆ", null, true));
            foreach (var paragraph in model.OtherTerms)
                body.Append(CreateParagraph(paragraph));

            var table = new Table();
            table.Append(new TableProperties(new TableBorders(
                new TopBorder { Val = BorderValues.None },
                new BottomBorder { Val = BorderValues.None },
                new LeftBorder { Val = BorderValues.None },
                new RightBorder { Val = BorderValues.None },
                new InsideHorizontalBorder { Val = BorderValues.None },
                new InsideVerticalBorder { Val = BorderValues.None }
            )));

            string[] lines = new[]
            {
            "(ลงชื่อ).................................................",
            "(                                   )",
            "สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม",
            "(ลงชื่อ).................................................",
            "(                                   )",
            model.AgencyName,
            "(ลงชื่อ).................................................",
            "(                                   )",
            "(ลงชื่อ).................................................",
            "(                                   )"
        };

            foreach (var line in lines)
            {
                var row = new TableRow();
                row.Append(new TableCell(new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = JustificationValues.Right },
                        new SpacingBetweenLines { After = "200" }),
                    new Run(
                        new RunProperties(
                            new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                            new FontSize { Val = "32" }),
                        new Text(line)))));
                table.Append(row);
            }

            body.Append(table);
            mainPart.Document.Save();
        }
        return stream.ToArray();
    }

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

    public byte[] GenImportContract(IEnumerable<ImportContractModels> items)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var package = new ExcelPackage();
        var ws = package.Workbook.Worksheets.Add("Contract");

        // ===== Header (Row 1) =====
        string[] headers = new[]
        {
        "ContractNumber","ProjectName","ContractParty","Domicile",
        "Start_Date","End_Date","Status","Amount","Installment","Owner",
        "ContractStorage","ContractType","InstallmentNo","PaymentDate","InstallmentAmount"
    };

        for (int i = 0; i < headers.Length; i++)
            ws.Cells[1, i + 1].Value = headers[i];

        using (var rng = ws.Cells[1, 1, 1, headers.Length])
        {
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            rng.Style.Font.Bold = true;
            rng.Style.Fill.PatternType = ExcelFillStyle.None;
        }

        // ===== Data Rows (starting Row 2) =====
        int r = 2;
        foreach (var x in items ?? Enumerable.Empty<ImportContractModels>())
        {
            ws.Cells[r, 1].Value = x.ContractNumber ?? "";
            ws.Cells[r, 2].Value = x.ProjectName ?? "";
            ws.Cells[r, 3].Value = x.ContractParty ?? "";
            ws.Cells[r, 4].Value = x.Domicile ?? "";

            // วันที่
            if (x.StartDate.HasValue) ws.Cells[r, 5].Value = x.StartDate.Value;
            if (x.EndDate.HasValue) ws.Cells[r, 6].Value = x.EndDate.Value;

            ws.Cells[r, 7].Value = x.Status ?? "";

            // ตัวเลข
            if (x.Amount.HasValue) ws.Cells[r, 8].Value = x.Amount.Value;
            if (x.Installment.HasValue) ws.Cells[r, 9].Value = x.Installment.Value;
            ws.Cells[r, 10].Value = x.Owner ?? "";

            // ชื่อคอลัมน์ในชีต = ContractStorage (ไม่มี underscore)
            ws.Cells[r, 11].Value = x.ContractStorage ?? "";

            ws.Cells[r, 12].Value = x.ContractType ?? "";
            if (x.InstallmentNo.HasValue) ws.Cells[r, 13].Value = x.InstallmentNo.Value;
            if (x.PaymentDate.HasValue) ws.Cells[r, 14].Value = x.PaymentDate.Value;
            if (x.InstallmentAmount.HasValue) ws.Cells[r, 15].Value = x.InstallmentAmount.Value;

            r++;
        }

        // ===== Number formats / Alignment =====
        ws.Column(5).Style.Numberformat.Format = "yyyy-mm-dd"; // Start_Date
        ws.Column(6).Style.Numberformat.Format = "yyyy-mm-dd"; // End_Date
        ws.Column(14).Style.Numberformat.Format = "yyyy-mm-dd"; // PaymentDate

        ws.Column(8).Style.Numberformat.Format = "#,##0.00";   // Amount
        ws.Column(15).Style.Numberformat.Format = "#,##0.00";   // InstallmentAmount

        ws.Column(9).Style.Numberformat.Format = "0";          // Installment
        ws.Column(13).Style.Numberformat.Format = "0";          // InstallmentNo

        // ตัวเลขจัดขวา
        ws.Column(8).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        ws.Column(9).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        ws.Column(13).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        ws.Column(15).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

        // Freeze header
        ws.View.FreezePanes(2, 1);

        // Autosize
        ws.Cells[ws.Dimension.Address].AutoFitColumns();

        return package.GetAsByteArray();
    }

}
