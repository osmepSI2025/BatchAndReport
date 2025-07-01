using BatchAndReport.Models;
using BatchAndReport.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Globalization;
using System.IO;
using System.Drawing;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;

public class WordWFService : IWordWFService
{
    public byte[] GenAnnualWorkProcesses()
    {
        using var stream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            body.Append(CreateHeading("การทบทวนกระบวนการของ ของ ฝผน. ประจำปี 2567", 20));
            body.Append(CreateNumberedParagraph("รายละเอียดประเด็นการทบทวน", new[]
            {
                    "ประกาศลักษณะงานฯ เรื่อง โครงสร้างและอัตรากำลัง การแบ่งส่วนงานและหน้าที่ความรับผิดชอบของหน่วยงานฯ (JD)",
                    "การพิจารณาระบบการควบคุมภายใน (คว.2)",
                    "การอ้างถึงงานปัจจุบัน และการได้ปรับบทบาทเพิ่มเติม"
                }));

            // ตารางเปรียบเทียบ 3 คอลัมน์
            body.Append(CreateThreeColumnTable());

            body.Append(CreateItalicNote("หมายเหตุ: *หมายหมาย JD/ **หมายหมาย คว.2/***หมายหมายการอ้างถึงงานปัจจุบัน"));

            body.Append(CreateBoldParagraph("กระบวนการที่จัดทำ Workflow เพิ่มเติม ได้แก่"));
            body.Append(CreateNormalParagraph("• การจัดทำแผนการส่งเสริม SME/นโยบาย/มาตรการ ให้กับหน่วยงานที่เกี่ยวข้อง*"));

            body.Append(CreateBoldParagraph("ความคิดเห็น"));
            body.Append(CreateCheckBoxOptions(new[]
            {
                    "เห็นชอบการปรับปรุง",
                    "มีความเห็นเพิ่มเติม"
                }));

            body.Append(CreateNormalParagraph("เห็นควรให้ตรวจสอบการศึกษาและจัดทำข้อมูลเพื่อใช้ในการส่งเสริม SME ยอด ..."));

            body.Append(CreateSignatureSection(
                leftName: "นางอธิศิณี ชาติธีระ", leftPosition: "รอ.ฝผส.",
                rightName: "นายธัชนะวัฒน์ โอภาสวัฒนา", rightPosition: "ผอ.ฝผส.",
                leftDate: "13 ก.ย. 67", rightDate: "13 ก.ย. 67"
            ));

            mainPart.Document.Save();
        }
        return stream.ToArray();
    }
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
    public byte[] GenWorkSystem()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var package = new ExcelPackage();
        var ws = package.Workbook.Worksheets.Add("AnnualProcess");

        // ===== Row 1 =====
        ws.Cells["A1:I1"].Merge = true;
        ws.Cells["A1"].Value = "กระบวนการทบทวนแบ่งตามกระบวนงาน";
        ws.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        ws.Cells["A1"].Style.Font.Bold = true;
        ws.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        ws.Cells["A1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 32, 96));
        ws.Cells["A1"].Style.Font.Color.SetColor(System.Drawing.Color.White);

        // ===== Row 2 =====
        ws.Cells["A2:C2"].Merge = true;
        ws.Cells["A2"].Value = "กระบวนการ (ทบทวน ปี 2567)";
        ws.Cells["D2"].Value = "2566";
        ws.Cells["E2"].Value = "ทบทวน";
        ws.Cells["F2"].Value = "หน่วยงาน";
        ws.Cells["G2"].Value = "Workflow";
        ws.Cells["H2"].Value = "WI";
        ws.Cells["I2"].Value = "ที่มา";
        ws.Cells["A2:I2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        ws.Cells["A2:I2"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 84, 127));
        ws.Cells["A2:I2"].Style.Font.Color.SetColor(System.Drawing.Color.White);
        ws.Cells["A2:I2"].Style.Font.Bold = true;

        // ===== Row 3 (Sub Header) =====
        ws.Cells["A3:I3"].Merge = true;
        ws.Cells["A3"].Value = "Core Process";
        ws.Cells["A3:I3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        ws.Cells["A3:I3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 112, 192));
        ws.Cells["A3:I3"].Style.Font.Color.SetColor(System.Drawing.Color.White);
        ws.Cells["A3:I3"].Style.Font.Bold = true;

        // ===== Row 4 (Column Headers) =====
        ws.Cells["A4"].Value = "No.";
        ws.Cells["B4"].Value = "C1";
        ws.Cells["C4"].Value = "การรวบรวมและวิเคราะห์ข้อมูล (BIG DATA)";
        ws.Cells["D4"].Value = "";
        ws.Cells["E4"].Value = "";
        ws.Cells["F4"].Value = "";
        ws.Cells["G4"].Value = "";
        ws.Cells["H4"].Value = "";
        ws.Cells["I4"].Value = "";

        ws.Cells["A4:I4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        ws.Cells["A4:I4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(221, 235, 247)); // ฟ้าอ่อน
        ws.Cells["A4:I4"].Style.Font.Bold = true;
        ws.Cells["A4:I4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

        // ===== Data Rows =====
        var data = GetTemplateData();
        int row = 5;
        foreach (var item in data)
        {
            ws.Cells[row, 1].Value = item.No;
            ws.Cells[row, 2].Value = item.Code;
            ws.Cells[row, 3].Value = item.Description;
            ws.Cells[row, 4].Value = item.Col2566;
            ws.Cells[row, 5].Value = item.ColReview;
            ws.Cells[row, 6].Value = item.Department;
            ws.Cells[row, 7].Value = item.Workflow;
            ws.Cells[row, 8].Value = item.WI;
            ws.Cells[row, 9].Value = item.Source;
            row++;
        }

        ws.Cells.AutoFitColumns();
        return package.GetAsByteArray();
    }

    private List<WorkSystemModels> GetTemplateData()
    {
        return new List<WorkSystemModels>
            {
                new WorkSystemModels { CoreProcess = "5 C1", No = "1", Code = "C1.1", Description = "การจัดทำรายงานสถานการณ์ MSME", Col2566 = "C1.1", ColReview = "C1.1", Department = "ฝผส.", Workflow = "มี workflow", WI = "มี WI", Source = "คว.2" },
                new WorkSystemModels { CoreProcess = "", No = "2", Code = "C1.2", Description = "การทบทวนตัวชี้วัดฯ", Col2566 = "C1.4", ColReview = "C1.4", Department = "ฝผพ.", Workflow = "มี workflow", WI = "มี WI", Source = "JD" },
                new WorkSystemModels { CoreProcess = "", No = "3", Code = "C1.3", Description = "การให้บริการทะเบียนและรับผู้ข้อมูลผู้ประกอบการ", Col2566 = "C1.6", ColReview = "C1.6", Department = "ฝผน.", Workflow = "มี workflow", WI = "มี WI", Source = "JD" },
                // เพิ่มได้อีกตามต้องการ
            };
    }
    private Paragraph CreateHeading(string text, int fontSize)
    {
        return new Paragraph(
            new Run(
                new RunProperties(new Bold(), new FontSize { Val = (fontSize * 2).ToString() }),
                new Text(text)
            )
        );
    }

    private Paragraph CreateNumberedParagraph(string title, string[] items)
    {
        var para = new Paragraph(new Run(new Bold(), new Text(title)));
        foreach (var item in items.Select((text, i) => $"{i + 1}. {text}"))
        {
            para.Append(new Run(new Break()), new Run(new Text(item)));
        }
        return para;
    }

    private Table CreateThreeColumnTable()
    {
        var table = new Table(new TableProperties(
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single },
                new BottomBorder { Val = BorderValues.Single },
                new LeftBorder { Val = BorderValues.Single },
                new RightBorder { Val = BorderValues.Single },
                new InsideHorizontalBorder { Val = BorderValues.Single },
                new InsideVerticalBorder { Val = BorderValues.Single }
            )));

        // Header
        table.Append(CreateTableRow("กระบวนการ ปี 2566 (เดิม)", "กระบวนการ ปี 2567 (ทบทวน)", "กิจกรรมควบคุม (Control Activity)"));

        // 5 รายการข้อมูล
        for (int i = 1; i <= 5; i++)
        {
            table.Append(CreateTableRow(
                $"การดำเนินงานเดิมข้อ {i}",
                $"ทบทวนปี 2567 ข้อ {i}",
                $"กิจกรรมควบคุมข้อ {i}"
            ));
        }

        return table;
    }

    private Paragraph CreateItalicNote(string text)
    {
        return new Paragraph(
            new Run(new RunProperties(new Italic()), new Text(text))
        );
    }

    private Paragraph CreateCheckBoxOptions(string[] options)
    {
        var para = new Paragraph();
        foreach (var opt in options)
        {
            para.Append(new Run(new Text("☐ " + opt)), new Run(new Break()));
        }
        return para;
    }

    private TableRow CreateTableRow(params string[] cells)
    {
        var row = new TableRow();
        foreach (var cellText in cells)
        {
            row.Append(new TableCell(new Paragraph(new Run(new Text(cellText)))));
        }
        return row;
    }

    private void AppendSignatureCell(Table table, string name, string position, string date)
    {
        var cell = new TableCell(
            new Paragraph(new Run(new Text(name))),
            new Paragraph(new Run(new Text(position))),
            new Paragraph(new Run(new Text("วันที่ " + date)))
        );
        var row = new TableRow();
        row.Append(cell);
        table.Append(row);
    }

    private Table CreateSignatureSection(string leftName, string leftPosition, string rightName, string rightPosition, string leftDate, string rightDate)
    {
        var table = new Table(new TableProperties(
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }
        ));

        var row1 = new TableRow();
        row1.Append(
            new TableCell(new Paragraph(new Run(new Text(leftName)))),
            new TableCell(new Paragraph(new Run(new Text(rightName))))
        );

        var row2 = new TableRow();
        row2.Append(
            new TableCell(new Paragraph(new Run(new Text(leftPosition)))),
            new TableCell(new Paragraph(new Run(new Text(rightPosition))))
        );

        var row3 = new TableRow();
        row3.Append(
            new TableCell(new Paragraph(new Run(new Text("วันที่ " + leftDate)))),
            new TableCell(new Paragraph(new Run(new Text("วันที่ " + rightDate))))
        );

        table.Append(row1, row2, row3);
        return table;
    }
    private Paragraph CreateBoldParagraph(string text)
    {
        return new Paragraph(
            new Run(
                new RunProperties(new Bold()),
                new Text(text)
            )
        );
    }
    private Paragraph CreateNormalParagraph(string text)
    {
        return new Paragraph(new Run(new Text(text ?? "")));
    }


}
