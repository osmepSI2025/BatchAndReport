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
using System.Xml;
using System.Xml.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

public class WordWFService : IWordWFService
{
    public byte[] GenAnnualWorkProcesses(WFProcessDetailModels detail)
    {
        using var stream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            var fiscalYear = detail.FiscalYear.ToString();
            var fiscalYearPrev = detail.FiscalYearPrevious;

            body.Append(CreateHeading($"การทบทวนกระบวนการของ ของ ฝผน. ประจำปี {fiscalYear}", 20));
            body.Append(CreateNumberedParagraph("รายละเอียดประเด็นการทบทวน", detail.ReviewDetails));
            body.Append(CreateEmptyLine());
            body.Append(CreateHeadingLeft($"การทบทวนกระบวนการของ ของ ฝผน. ประจำปี {fiscalYear} ดังนี้", 20));
            body.Append(CreateThreeColumnTable(fiscalYearPrev, fiscalYear, detail.PrevProcesses, detail.CurrentProcesses, detail.ControlActivities));

            body.Append(CreateItalicNote("หมายเหตุ: *หมายหมาย JD/ **หมายหมาย คว.2/***หมายหมายการอ้างถึงงานปัจจุบัน"));

            body.Append(CreateEmptyLine());

            body.Append(CreateBoldParagraph("กระบวนการที่จัดทำ Workflow เพิ่มเติม ได้แก่", 20));
            foreach (var wf in detail.WorkflowProcesses)
                body.Append(CreateNormalParagraph($"• {wf}"));

            body.Append(CreateBoldParagraph("ความคิดเห็น", 20));
            body.Append(CreateCheckBoxOptions(new[] {
            "เห็นชอบการปรับปรุง",
            "มีความเห็นเพิ่มเติม"
        }));
            body.Append(CreateEmptyLine());
            body.Append(CreateEmptyLine());
            foreach (var r in detail.ApproveRemarks)
                body.Append(CreateNormalParagraph(r));

            body.Append(CreateSignatureSection(
                leftName: detail.Approver1Name ?? "(ชื่อผู้ลงนาม 1)", leftPosition: detail.Approver1Position ?? "ตำแหน่ง",
                rightName: detail.Approver2Name ?? "(ชื่อผู้ลงนาม 2)", rightPosition: detail.Approver2Position ?? "ตำแหน่ง",
                leftDate: detail.Approve1Date ?? "ไม่พบข้อมูล", rightDate: detail.Approve2Date ?? "ไม่พบข้อมูล"
            , TableRowAlignmentValues.Center));

            mainPart.Document.Save();
        }
        return stream.ToArray();
    }
    public byte[] ConvertWordToPdf(byte[] wordBytes)
    {
        try
        {
            using var inputStream = new MemoryStream(wordBytes);
            var doc = new Spire.Doc.Document();
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
    public byte[] GenWorkSystem(WorkSystemModels model)
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
        // ✅ เพิ่มตรงนี้
        int startRow = 5;
        int no = 1;
        foreach (var item in model.ProcessDetails)
        {
            ws.Cells[startRow, 1].Value = no;
            ws.Cells[startRow, 2].Value = item.ProcessCode;
            ws.Cells[startRow, 3].Value = item.ProcessName;
            ws.Cells[startRow, 4].Value = item.PrevProcessCode;
            ws.Cells[startRow, 5].Value = item.ReviewType;
            ws.Cells[startRow, 6].Value = item.Department;
            ws.Cells[startRow, 7].Value = item.Workflow;
            ws.Cells[startRow, 8].Value = item.WI;
            ws.Cells[startRow, 9].Value = ""; // ที่มา
            startRow++;
            no++;
        }


        ws.Cells.AutoFitColumns();
        return package.GetAsByteArray();
    }
    public byte[] GenInternalControlSystem(List<WFInternalControlProcessModels> detail)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var package = new ExcelPackage();
        var ws = package.Workbook.Worksheets.Add("ImportantProcess");

        // ===== Header =====
        ws.Cells["A1:C1"].Merge = true;
        ws.Cells["A1"].Value = "กระบวนการทำงานที่สำคัญตามรายงานการจัดวางระบบการควบคุมภายใน";
        StyleHeader(ws.Cells["A1"], bold: true);

        ws.Cells["A2"].Value = "ภารกิจตามกฎหมายที่จัดตั้งหน่วยงานของรัฐหรือตามแผนการดำเนินงานหรืองานอื่นๆ ที่สำคัญ";
        ws.Cells["B2"].Value = "ลำดับ";
        ws.Cells["C2"].Value = "ของหน่วยงานของรัฐ/วัตถุประสงค์";
        StyleHeader(ws.Cells["A2"], bold: true);
        StyleHeader(ws.Cells["B2"], bold: true);
        StyleHeader(ws.Cells["C2"], bold: true);

        // ===== Green Bar =====
        ws.Cells["A3:C3"].Merge = true;
        ws.Cells["A3"].Value = "งานนโยบายและยุทธศาสตร์";
        ws.Cells["A3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        ws.Cells["A3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(198, 224, 180));
        ws.Cells["A3"].Style.Font.Bold = true;

        // ===== Fill Data =====
        int startRow = 4;
        int index = 1;

        foreach (var item in detail)
        {
            // A: ชื่อแผน + วัตถุประสงค์
            ws.Cells[$"A{startRow}"].Value = $"{item.PlanCategoryName}\n\nวัตถุประสงค์: {item.Objective}";
            ws.Cells[$"A{startRow}"].Style.WrapText = true;
            ws.Cells[$"A{startRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

            // B: ลำดับ
            ws.Cells[$"B{startRow}"].Value = index++;

            // C: BusinessUnitId
            ws.Cells[$"C{startRow}"].Value = item.BusinessUnitId;

            startRow++;
        }

        // ===== Style B3:C3 =====
        using (var range = ws.Cells["B3:C3"])
        {
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 84, 127));
            range.Style.Font.Color.SetColor(System.Drawing.Color.White);
            range.Style.Font.Bold = true;
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        // ===== Border =====
        var usedRange = ws.Cells[$"A1:C{startRow - 1}"];
        usedRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
        usedRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
        usedRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        usedRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

        ws.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        ws.Cells.AutoFitColumns();

        return package.GetAsByteArray();
    }
    public async Task<byte[]> GenInternalControlSystemWord(List<WFInternalControlProcessModels> detail, WFSubProcessDetailModels detail2)
    {
        using var stream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            // === Section 1: Internal Control System Table ===
            body.Append(CreateHeading("กระบวนการทำงานที่สำคัญตามรายงานการจัดวางระบบการควบคุมภายใน", 22));
            body.Append(CreateBoldParagraph("ภารกิจตามกฎหมายที่จัดตั้งหน่วยงานของรัฐหรือตามแผนการดำเนินงานหรืองานอื่นๆ ที่สำคัญ", 22));
            body.Append(CreateTableHeader());

            var greenRow = new TableRow(
                CreateMergedCell("งานนโยบายและยุทธศาสตร์", 3, "C6EFCE", bold: true)
            );

            var table = CreateFullWidthTable();
            table.Append(greenRow);

            int index = 1;
            foreach (var item in detail)
            {
                var planCell = CreateCell(item.PlanCategoryName + "\n\nวัตถุประสงค์: " + item.Objective, JustificationValues.Left, wrap: true);
                var indexCell = CreateCell(index.ToString(), JustificationValues.Center);
                var buCell = CreateCell(item.BusinessUnitId ?? "-", JustificationValues.Center);
                var row = new TableRow();
                row.Append(planCell, indexCell, buCell);
                table.Append(row);
                index++;
            }

            body.Append(table);
            body.Append(CreateEmptyLine());

            // === Section Break + Header ===
            await AddDocumentHeader(mainPart, detail2);

            body = mainPart.Document.AppendChild(new Body());

            // === Section 2: Evaluation ===
            var evals = detail2.Evaluations?.Select(e => e.EvaluationDesc).Where(e => !string.IsNullOrEmpty(e)).ToArray();
            body.Append(CreateBoldParagraph("ตัวชี้วัดของกระบวนการ :", 20));
            body.Append(evals?.Length > 0 ? (OpenXmlElement)CreateNumberedList(evals) : CreateNormalParagraph("-"));
            body.Append(CreateEmptyLine());

            // === Section 3: Approvals ===
            body.Append(CreateApprovalsTable(detail2.Approvals));
            body.Append(CreateEmptyLine());

            // === Section 4: Revisions ===
            var revTable = CreateFullWidthTable();
            revTable.Append(new TableRow(new[]
            {
                new TableCell(
                    new TableCellProperties(
                        new GridSpan { Val = 3 },
                        new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
                        new Shading { Fill = "D9D9D9", Val = ShadingPatternValues.Clear, Color = "auto" },
                        new TableCellBorders(
                            new TopBorder { Val = BorderValues.Single },
                            new BottomBorder { Val = BorderValues.Single },
                            new LeftBorder { Val = BorderValues.Single },
                            new RightBorder { Val = BorderValues.Single }
                        )
                    ),
                    new Paragraph(
                        new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                        new Run(new Text("ประวัติการแก้ไขเอกสาร"))
                    )
                )
            }));

            revTable.Append(new TableRow(new[]
            {
                CreateCell("ครั้งที่แก้ไข", JustificationValues.Center),
                CreateCell("วันที่แก้ไข", JustificationValues.Center),
                CreateCell("รายละเอียด", JustificationValues.Center)
            }));

            if (detail2.Revisions?.Count > 0)
            {
                int i = 1;
                foreach (var rev in detail2.Revisions)
                {
                    revTable.Append(new TableRow(new[]
                    {
                        CreateCell(i++.ToString(), JustificationValues.Center),
                        CreateCell(rev.DateTime?.ToString("d MMM yy", new CultureInfo("th-TH")) ?? "-", JustificationValues.Center),
                        CreateCell(rev.EditDetail ?? "-", JustificationValues.Left)
                    }));
                }
            }
            else
            {
                revTable.Append(new TableRow(new[]
                {
                    CreateCell("-", JustificationValues.Center),
                    CreateCell("-", JustificationValues.Center),
                    CreateCell("-", JustificationValues.Center)
                }));
            }

            body.Append(revTable);
            body.Append(CreateEmptyLine());

            // === Section 5: Control Points ===
            var cpTable = CreateFullWidthTable();
            cpTable.Append(new TableRow(new[]
            {
                CreateCell("จุดควบคุม", JustificationValues.Center),
                CreateCell("กิจกรรมควบคุม", JustificationValues.Center),
                CreateCell("รายละเอียด", JustificationValues.Center)
            }));

            if (detail2.ControlPoints?.Count > 0)
            {
                foreach (var cp in detail2.ControlPoints)
                {
                    cpTable.Append(new TableRow(new[]
                    {
                        CreateCell(cp.ProcessControlCode ?? "-", JustificationValues.Center),
                        CreateCell(cp.ProcessControlActivity ?? "-", JustificationValues.Center),
                        CreateCell(cp.ProcessControlDetail ?? "-", JustificationValues.Center)
                    }));
                }
            }
            else
            {
                cpTable.Append(new TableRow(new[]
                {
                    CreateCell("-", JustificationValues.Center),
                    CreateCell("-", JustificationValues.Center),
                    CreateCell("-", JustificationValues.Center)
                }));
            }

            body.Append(cpTable);
            body.Append(CreateEmptyLine());

            // === Section 6: Diagrams ===
            if (!string.IsNullOrEmpty(detail2.Header?.DiagramAttachFile))
            {
                var base64 = ExtractBase64FromXml(detail2.Header.DiagramAttachFile);
                if (!string.IsNullOrEmpty(base64))
                {
                    try
                    {
                        byte[] imageBytes = Convert.FromBase64String(base64);
                        AddDiagramImagePage(body, imageBytes, mainPart);
                    }
                    catch (FormatException)
                    {
                        Console.WriteLine("Invalid base64 DiagramAttachFile.");
                    }
                }
            }

            if (!string.IsNullOrEmpty(detail2.Header?.ProcessAttachFile))
            {
                var base64 = ExtractBase64FromXml(detail2.Header.ProcessAttachFile);
                if (!string.IsNullOrEmpty(base64))
                {
                    try
                    {
                        byte[] imageBytes = Convert.FromBase64String(base64);
                        AddDiagramImagePage(body, imageBytes, mainPart);
                    }
                    catch (FormatException)
                    {
                        Console.WriteLine("Invalid base64 ProcessAttachFile.");
                    }
                }
            }

            mainPart.Document.Save();
        }

        return stream.ToArray();
    }
    public async Task<byte[]> GenWorkProcessPoint(WFSubProcessDetailModels detail)
    {
        using var stream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            mainPart.Document.Append(body);
            // Add Header First
            await AddDocumentHeader(mainPart, detail);

            body = mainPart.Document.AppendChild(new Body());

            var lastRevDate = detail.Revisions?.LastOrDefault()?.DateTime;
            var revDateText = lastRevDate.HasValue ? lastRevDate.Value.ToString("dd MMM yy", new CultureInfo("th-TH")) : "-";

            // Evaluation
            var evals = detail.Evaluations?.Select(e => e.EvaluationDesc).Where(e => !string.IsNullOrEmpty(e)).ToArray();
            body.Append(CreateBoldParagraph("ตัวชี้วัดของกระบวนการ :", 20));
            if (evals?.Length > 0)
                body.Append(CreateNumberedList(evals));
            else
                body.Append(CreateNormalParagraph("-"));
            body.Append(CreateEmptyLine());

            // Approvals Table
            body.Append(CreateApprovalsTable(detail.Approvals));
            body.Append(CreateEmptyLine());

            // Revisions Table
            var revTable = CreateFullWidthTable();
            revTable.Append(new TableRow(new[]
            {
            new TableCell(
                new TableCellProperties(
                    new GridSpan { Val = 3 },
                    new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
                    new Shading { Fill = "D9D9D9", Val = ShadingPatternValues.Clear, Color = "auto" },
                    new TableCellBorders(
                        new TopBorder { Val = BorderValues.Single },
                        new BottomBorder { Val = BorderValues.Single },
                        new LeftBorder { Val = BorderValues.Single },
                        new RightBorder { Val = BorderValues.Single }
                    )
                ),
                new Paragraph(
                    new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                    new Run(new Text("ประวัติการแก้ไขเอกสาร"))
                )
            )
        }));

            revTable.Append(new TableRow(new[]
            {
            CreateCell("ครั้งที่แก้ไข", JustificationValues.Center),
            CreateCell("วันที่แก้ไข", JustificationValues.Center),
            CreateCell("รายละเอียด", JustificationValues.Center)
        }));

            if (detail.Revisions?.Count > 0)
            {
                int i = 1;
                foreach (var rev in detail.Revisions)
                {
                    revTable.Append(new TableRow(new[]
                    {
                    CreateCell(i++.ToString(), JustificationValues.Center),
                    CreateCell(rev.DateTime?.ToString("d MMM yy", new CultureInfo("th-TH")) ?? "-", JustificationValues.Center),
                    CreateCell(rev.EditDetail ?? "-", JustificationValues.Left)
                }));
                }
            }
            else
            {
                revTable.Append(new TableRow(new[]
                {
                CreateCell("-", JustificationValues.Center),
                CreateCell("-", JustificationValues.Center),
                CreateCell("-", JustificationValues.Center)
            }));
            }

            body.Append(revTable);
            body.Append(CreateEmptyLine());

            // Control Points
            var cpTable = CreateFullWidthTable();
            cpTable.Append(new TableRow(new[]
            {
            CreateCell("จุดควบคุม", JustificationValues.Center),
            CreateCell("กิจกรรมควบคุม", JustificationValues.Center),
            CreateCell("รายละเอียด", JustificationValues.Center)
        }));

            if (detail.ControlPoints?.Count > 0)
            {
                foreach (var cp in detail.ControlPoints)
                {
                    cpTable.Append(new TableRow(new[]
                    {
                    CreateCell(cp.ProcessControlCode ?? "-", JustificationValues.Center),
                    CreateCell(cp.ProcessControlActivity ?? "-", JustificationValues.Center),
                    CreateCell(cp.ProcessControlDetail ?? "-", JustificationValues.Center)
                }));
                }
            }
            else
            {
                cpTable.Append(new TableRow(new[]
                {
                CreateCell("-", JustificationValues.Center),
                CreateCell("-", JustificationValues.Center),
                CreateCell("-", JustificationValues.Center)
            }));
            }

            body.Append(cpTable);

            if (!string.IsNullOrEmpty(detail.Header?.DiagramAttachFile))
            {
                var base64 = ExtractBase64FromXml(detail.Header.DiagramAttachFile);
                if (!string.IsNullOrEmpty(base64))
                {
                    try
                    {
                        byte[] imageBytes = Convert.FromBase64String(base64);
                        AddDiagramImagePage(body, imageBytes, mainPart);
                    }
                    catch (FormatException)
                    {
                        Console.WriteLine("Invalid base64 DiagramAttachFile.");
                    }
                }
            }

            if (!string.IsNullOrEmpty(detail.Header?.ProcessAttachFile))
            {
                var base64 = ExtractBase64FromXml(detail.Header.ProcessAttachFile);
                if (!string.IsNullOrEmpty(base64))
                {
                    try
                    {
                        byte[] imageBytes = Convert.FromBase64String(base64);
                        AddDiagramImagePage(body, imageBytes, mainPart);
                    }
                    catch (FormatException)
                    {
                        Console.WriteLine("Invalid base64 ProcessAttachFile.");
                    }
                }
            }


            mainPart.Document.Save();
        }

        return stream.ToArray();
    }

    public List<string> ExtractFilePathsFromXml(string? xml)
    {
        var filePaths = new List<string>();
        if (string.IsNullOrWhiteSpace(xml)) return filePaths;

        try
        {
            var xdoc = XDocument.Parse(xml);
            var fields = xdoc.Descendants("field")
                             .Where(f => (string?)f.Attribute("name") == "FilePath")
                             .Select(f => f.Element("value")?.Value)
                             .Where(v => !string.IsNullOrWhiteSpace(v));

            filePaths.AddRange(fields!);
        }
        catch (Exception ex)
        {
            Console.WriteLine("❌ XML parse error: " + ex.Message);
        }

        return filePaths;
    }
    private string? ExtractBase64FromXml(string xml)
    {
        try
        {
            var doc = new XmlDocument();
            doc.LoadXml(xml);
            var node = doc.SelectSingleNode("//content");
            return node?.InnerText?.Trim();
        }
        catch
        {
            return null;
        }
    }

    public byte[] GenWFProcessDetail(WFProcessDetailModels detail)
    {
        using var stream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            // Header
            body.Append(CreateBoldParagraph(
                $"การทบทวนกลุ่มกระบวนการหลักและกลุ่มกระบวนการสนับสนุน ประจำปีงบประมาณ {detail.FiscalYear}", 20));
            body.Append(CreateEmptyLine());

            // === Core Process Table ===
            var coreTable = CreateFullWidthTable();

            // Row 1: กลุ่มกระบวนการหลัก + รหัส
            var coreRow1 = new TableRow();
            coreRow1.Append(CreateCellFixed("กลุ่มกระบวนการหลัก\n(Core Process)", JustificationValues.Left, "10%", alignTop: true, rowspan: 2));

            foreach (var core in detail.CoreProcesses)
            {
                coreRow1.Append(CreateCellFixed(core.ProcessGroupCode, JustificationValues.Center, "10%", alignTop: true, shading: "00C896"));
            }
            coreTable.Append(coreRow1);

            // Row 2: ชื่อกระบวนการ
            var coreRow2 = new TableRow();
            coreRow2.Append(CreateMergedEmptyCell()); // สำหรับ rowspan
            foreach (var core in detail.CoreProcesses)
            {
                coreRow2.Append(CreateCellFixed(core.ProcessGroupName, JustificationValues.Center, "10%", alignTop: true, shading: "00C896"));
            }
            coreTable.Append(coreRow2);
            body.Append(coreTable);
            body.Append(CreateEmptyLine());

            // === Supporting Process Table ===
            var supportTable = CreateFullWidthTable();
            int supportCount = detail.SupportProcesses.Count;

            for (int i = 0; i < supportCount; i++)
            {
                var support = detail.SupportProcesses[i];
                var row = new TableRow();

                // คอลัมน์ 1: Vertical Merge
                if (i == 0)
                {
                    row.Append(CreateCellFixed("กลุ่มกระบวนการสนับสนุน\n(Supporting Process)",
                        JustificationValues.Left, "10%", alignTop: true, verticalMerge: "restart"));
                }
                else
                {
                    row.Append(CreateCellFixed("", JustificationValues.Left, "10%", alignTop: true, verticalMerge: "continue"));
                }

                // คอลัมน์ 2-3
                row.Append(CreateCellFixed(support.ProcessGroupCode, JustificationValues.Center, "10%", alignTop: true, shading: "4CB1F0"));
                row.Append(CreateCellFixed(support.ProcessGroupName, JustificationValues.Left, "80%", alignTop: true, shading: "4CB1F0"));

                supportTable.Append(row);
            }

            body.Append(supportTable);
            body.Append(CreateEmptyLine());

            mainPart.Document.Save();
        }

        return stream.ToArray();
    }

    public byte[] GenWorkProcessPointPreview()
    {
        // Mock data
        var workflow = new WorkflowPoint
        {
            WorkflowTitle = "C2.1 การจัดทำแผนการส่งเสริม SMEs",
            Department = "ฝ่ายนโยบายและแผนการส่งเสริม SMEs (ฝผย.)",
            Indicators = "จัดประชุม เพื่อจัดทำแผนการส่งเสริม SMEs||มีแผนการส่งเสริม SMEs ระยะ 5 ปี",
            EditNumber = 2,
            EditDate = new DateTime(2025, 11, 22),
            PageNumber = "1/5",
            Approvals = new List<WorkflowApproval>
            {
                new WorkflowApproval { FullName = "นายสุปรีย์ เถระพันธ์", Position = "หัวหน้าฝ่ายนโยบายและแผนการส่งเสริม SMEs", SignText = "ลงนาม", Level = 1 },
                new WorkflowApproval { FullName = "นางสาวอัญชรินธร จิรโชติวิศาลพันธ์", Position = "รองผู้อำนวยการ ฝ่ายนโยบายและแผนการส่งเสริม SMEs", SignText = "ลงนาม", Level = 2 },
                new WorkflowApproval { FullName = "นายธัชนะวัฒน์ โอภาสวัฒนา", Position = "ผู้อำนวยการ ฝ่ายนโยบายและแผนการส่งเสริม SMEs", SignText = "ลงนาม", Level = 3 }
            },
            HistoryEdits = new List<WorkflowHistory>
            {
                new WorkflowHistory { EditNumber = 1, EditDate = new DateTime(2024, 10, 3), Description = "ปรับปรุง Control Point และรายละเอียดขั้นตอน" },
                new WorkflowHistory { EditNumber = 2, EditDate = new DateTime(2025, 11, 22), Description = "เพิ่มรายละเอียดความร่วมมือกับหน่วยงานภายนอก" }
            }
        };

        using var stream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            body.Append(CreateHeaderBoxTable(workflow));
            body.Append(CreateBoldParagraph("ขั้นตอนการปฏิบัติงาน (Workflow)", 24));
            body.Append(CreateEmptyLine());
            body.Append(CreateNormalParagraph(workflow.WorkflowTitle));
            body.Append(CreateNormalParagraph("หน่วยงาน: " + workflow.Department));
            body.Append(CreateEmptyLine());
            body.Append(CreateNumberedList(workflow.Indicators.Split("||")));
            body.Append(CreateEmptyLine());
            body.Append(CreateApprovalTable(workflow.Approvals.ToList()));
            body.Append(CreateEmptyLine());
            body.Append(CreateBoldParagraph("ประวัติการแก้ไขเอกสาร", 20));
            body.Append(CreateHistoryTable(workflow.HistoryEdits.ToList()));
            body.Append(CreateBoldParagraph("สรุปประเด็นเพิ่มเติม", 20));
            body.Append(CreateNormalParagraph("มีการจัดทำ Workflow เพิ่มเติม ในขั้นตอนที่เกี่ยวข้องกับการจัดทำแผนการส่งเสริม SMEs ให้สอดคล้องกับบทบาทของหน่วยงาน"));
            body.Append(CreateNormalParagraph("แผนดังกล่าวมีผลกระทบกับการจัดสรรงบประมาณและความร่วมมือกับหน่วยงานอื่น ๆ"));
            body.Append(CreateNormalParagraph("ควรมีการปรับปรุง Control Point ให้มีความชัดเจนและเป็นรูปธรรม"));

            mainPart.Document.Save();
        }
        return stream.ToArray();
    }

    private Table CreateTableHeader()
    {
        var table = CreateFullWidthTable();
        var headerRow = new TableRow(new[]
        {
        CreateCell("ของหน่วยงานของรัฐ/วัตถุประสงค์", JustificationValues.Center, bold: true),
        CreateCell("ลำดับ", JustificationValues.Center, bold: true),
        CreateCell("หน่วยงาน", JustificationValues.Center, bold: true),
    });
        table.Append(headerRow);
        return table;
    }
    private TableCell CreateCell(string text, JustificationValues alignment, bool bold = false, bool wrap = false)
    {
        var cellProps = new TableCellProperties();
        if (wrap)
        {
            cellProps.Append(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Top });
        }

        var para = new Paragraph(
            new ParagraphProperties(new Justification { Val = alignment }),
            new Run(
                new RunProperties
                {
                    Bold = bold ? new Bold() : null,
                    FontSize = new FontSize { Val = "20" }, // ~20pt
                    RunFonts = new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" }
                },
                new Text(text) { Space = SpaceProcessingModeValues.Preserve }
            )
        );

        return new TableCell(cellProps, para);
    }

    private TableCell CreateMergedCell(string text, int colspan, string bgColorHex, bool bold = false)
    {
        return new TableCell(
            new TableCellProperties(
                new GridSpan { Val = colspan },
                new Shading { Fill = bgColorHex, Color = "auto", Val = ShadingPatternValues.Clear },
                new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "100%" }
            ),
            new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
                new Run(
                    new RunProperties
                    {
                        Bold = bold ? new Bold() : null,
                        FontSize = new FontSize { Val = "20" },
                        RunFonts = new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" }
                    },
                    new Text(text)
                )
            )
        );
    }
    private TableCell CreateCellFixed(string text, JustificationValues align, string widthPercent, bool alignTop = false, string shading = null, int rowspan = 1, string verticalMerge = null)
    {
        var props = new TableCellProperties(
            new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = widthPercent },
            new TableCellMarginDefault(
                new TopMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                new BottomMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                new LeftMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                new RightMargin { Width = "100", Type = TableWidthUnitValues.Dxa }
            )
        );

        if (alignTop)
        {
            props.Append(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Top });
        }

        if (!string.IsNullOrEmpty(shading))
        {
            props.Append(new Shading { Fill = shading, Val = ShadingPatternValues.Clear, Color = "auto" });
        }

        if (rowspan > 1)
        {
            props.Append(new VerticalMerge { Val = MergedCellValues.Restart });
            props.Append(new GridSpan { Val = 1 });
        }

        if (!string.IsNullOrEmpty(verticalMerge))
        {
            props.Append(new VerticalMerge { Val = verticalMerge == "restart" ? MergedCellValues.Restart : MergedCellValues.Continue });
        }

        return new TableCell(
            props,
            new Paragraph(
                new ParagraphProperties(new Justification { Val = align }),
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                        new FontSize { Val = "20" }
                    ),
                    new Text(text ?? "-") { Space = SpaceProcessingModeValues.Preserve }
                )
            )
        );
    }
    private static TableCell CreateCell(string text, JustificationValues align, int rowspan = 1, int colspan = 1, string? shading = null, string? verticalMerge = null)
    {
        // Ensure all code paths return a value
        var cellProperties = new TableCellProperties();

        if (rowspan > 1)
        {
            cellProperties.Append(new VerticalMerge { Val = MergedCellValues.Restart });
        }

        if (colspan > 1)
        {
            cellProperties.Append(new GridSpan { Val = colspan });
        }

        if (!string.IsNullOrEmpty(shading))
        {
            cellProperties.Append(new Shading
            {
                Fill = shading,
                Val = ShadingPatternValues.Clear,
                Color = "auto"
            });
        }

        if (!string.IsNullOrEmpty(verticalMerge))
        {
            cellProperties.Append(new VerticalMerge { Val = verticalMerge == "restart" ? MergedCellValues.Restart : MergedCellValues.Continue });
        }

        var paragraph = new Paragraph(
            new ParagraphProperties(new Justification { Val = align }),
            new Run(new Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve })
        );

        return new TableCell(cellProperties, paragraph);
    }
    private TableCell CreateMergedEmptyCell()
    {
        return new TableCell(
            new TableCellProperties(
                new VerticalMerge { Val = MergedCellValues.Continue }),
            new Paragraph());
    }
    private void AddDiagramImagePage(Body body, byte[] imageBytes, MainDocumentPart mainPart)
    {
        var imagePart = mainPart.AddImagePart(ImagePartType.Png); // หรือใช้ ImagePartType.Jpeg ถ้าเป็น .jpg
        using var stream = new MemoryStream(imageBytes);
        imagePart.FeedData(stream);

        var imageId = mainPart.GetIdOfPart(imagePart);

        var element = new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = 5000000, Cy = 4000000 },
                new DW.EffectExtent
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DW.DocProperties
                {
                    Id = (UInt32Value)1U,
                    Name = "Diagram"
                },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = (UInt32Value)0U, Name = "Diagram.png" },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(
                            new A.Blip { Embed = imageId },
                            new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = 5000000, Cy = 4000000 }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
                )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            )
            {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)0U,
                DistanceFromRight = (UInt32Value)0U,
            });

        // Page break ก่อนรูป
        body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
        body.Append(new Paragraph(new Run(element)));
    }

    private async Task AddDocumentHeader(MainDocumentPart mainPart, WFSubProcessDetailModels detail)
    {
        var headerPart = mainPart.AddNewPart<HeaderPart>();
        string headerPartId = mainPart.GetIdOfPart(headerPart);

        var header = new Header();
        var table = new Table();

        // ➤ Table Properties (เต็มความกว้าง + เส้นกรอบ)
        table.AppendChild(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single },
                new BottomBorder { Val = BorderValues.Single },
                new LeftBorder { Val = BorderValues.Single },
                new RightBorder { Val = BorderValues.Single },
                new InsideHorizontalBorder { Val = BorderValues.Single },
                new InsideVerticalBorder { Val = BorderValues.Single }
            )
        ));

        // 🔹 โหลดโลโก้จาก local path (แทนที่ HttpClient)
        byte[] logoBytes = null;
        var imagePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
        //if (File.Exists(imagePath))
        //{
        //    logoBytes = await File.ReadAllBytesAsync(imagePath);
        //}

        //string imageId = null;
        //if (logoBytes != null)
        //{
        //    var imagePart = mainPart.AddImagePart(ImagePartType.Png);
        //    using (var stream = new MemoryStream(logoBytes))
        //        imagePart.FeedData(stream);
        //    imageId = mainPart.GetIdOfPart(imagePart);
        //}
        // Add image part and feed image data
        var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg, "rIdLogo");
        using (var imgStream = File.OpenRead(imagePath))
        {
            imagePart.FeedData(imgStream);
        }

        // 🔹 Row 1: โลโก้ + ข้อความ + แก้ไขล่าสุด
        var row1 = new TableRow();

        var leftCell = new TableCell(
            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "66%" })
        );

        if (imagePart != null)
        {
            leftCell.Append(
                new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
                    new Run(WordServiceSetting.CreateImage(
                     mainPart.GetIdOfPart(imagePart),
                     240, 80))) // ✅ เรียกเหมือน OnGetWordContact_SupportSMEsService
                    );
        }

        leftCell.Append(
            new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Right }),
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                        new FontSize { Val = "20" }, new Bold()
                    ),
                    new Text("ขั้นตอนการปฏิบัติงาน (Workflow)")
                )
            )
        );
        // ➤ จัดข้อความแนวตั้งชิดบน
        if (leftCell.TableCellProperties == null)
            leftCell.TableCellProperties = new TableCellProperties();

        leftCell.TableCellProperties.Append(
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Top }
        );

        // ✅ Right cell: แสดงข้อมูลการแก้ไข
        var rightCell = new TableCell(
            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "34%" })
        );

        var lastRev = detail.Revisions?.LastOrDefault()?.DateTime;
        var revDateText = lastRev.HasValue ? lastRev.Value.ToString("d MMM yy", new CultureInfo("th-TH")) : "-";

        rightCell.Append(
            new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
                new Run(new Text("ครั้งที่แก้ไข: " + (detail.Revisions?.Count ?? 0)))),
            new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
                new Run(new Text("วันที่แก้ไข: " + revDateText))),
            new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
                new Run(new Text("หน้า: 1/5")))
        );

        row1.Append(leftCell, rightCell);
        table.Append(row1);

        // 🔹 Row 2: ProcessCode + ProcessName
        table.Append(new TableRow(new TableCell(
            new TableCellProperties(
                new GridSpan { Val = 3 },
                new Shading { Fill = "DDEBF7", Val = ShadingPatternValues.Clear, Color = "auto" }
            ),
            new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
                new Run(new Text($"{detail.Header?.ProcessCode ?? "-"} {detail.Header?.ProcessName ?? "-"}"))
            )
        )));

        // 🔹 Row 3: Owner Business Unit
        table.Append(new TableRow(new TableCell(
            new TableCellProperties(new GridSpan { Val = 3 }),
            new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
                new Run(new Text("หน่วยงาน: " + (detail.OwnerBusinessUnitName ?? "-")))
            )
        )));

        // Append table to header
        header.Append(table);
        headerPart.Header = header;

        // Section Properties: Attach header
        var sectionProps = new SectionProperties();
        sectionProps.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = headerPartId });

        if (mainPart.Document.Body == null)
            mainPart.Document.AppendChild(new Body());

        mainPart.Document.Body.Append(sectionProps);
    }
    private Table CreateApprovalsTable(List<SubProcessReviewApproval> approvals)
    {
        var table = CreateFullWidthTable();

        // Row 1: Title row with merged cell
        var titleRow = new TableRow();
        titleRow.Append(new TableCell(
            new TableCellProperties(
                new GridSpan { Val = 4 },
                new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "10000" },
                new Shading { Fill = "D9D9D9", Val = ShadingPatternValues.Clear, Color = "auto" }
            ),
            new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                new Run(new Text("การอนุมัติเอกสาร"))
            )
        ));
        table.Append(titleRow);

        // Row 2: Header
        var headerRow = new TableRow(new[] {
        CreateCell("", JustificationValues.Center),
        CreateCell("ผู้จัดทำ", JustificationValues.Center),
        CreateCell("ผู้ตรวจสอบ", JustificationValues.Center),
        CreateCell("ผู้อนุมัติ", JustificationValues.Center)
    });
        table.Append(headerRow);

        // Row 3: ลงนาม
        var signRow = new TableRow();
        signRow.Append(CreateCell("ลงนาม", JustificationValues.Center));
        for (int i = 0; i < 3; i++)
        {
            var item = approvals.ElementAtOrDefault(i);
            signRow.Append(CreateCell(item != null ? "(ลายเซ็น)" : "-", JustificationValues.Center)); // Placeholder signature
        }
        table.Append(signRow);

        // Row 4: ชื่อ
        var nameRow = new TableRow();
        nameRow.Append(CreateCell("", JustificationValues.Center));
        for (int i = 0; i < 3; i++)
        {
            var item = approvals.ElementAtOrDefault(i);
            nameRow.Append(CreateCell(item != null ? $"({item.EmployeeId ?? "-"})" : "-", JustificationValues.Center));
        }
        table.Append(nameRow);

        // Row 5: ตำแหน่ง
        var posRow = new TableRow();
        posRow.Append(CreateCell("ตำแหน่ง", JustificationValues.Center));
        for (int i = 0; i < 3; i++)
        {
            var item = approvals.ElementAtOrDefault(i);
            posRow.Append(CreateCell(item?.ApprovalTypeCode ?? "-", JustificationValues.Center));
        }
        table.Append(posRow);

        return table;
    }
    private Table CreateFullWidthTable()
    {
        return new Table(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single },
                new BottomBorder { Val = BorderValues.Single },
                new LeftBorder { Val = BorderValues.Single },
                new RightBorder { Val = BorderValues.Single },
                new InsideHorizontalBorder { Val = BorderValues.Single },
                new InsideVerticalBorder { Val = BorderValues.Single }
            )));
    }
    private TableCell CreateCell(string text, JustificationValues align)
    {
        return new TableCell(
            new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Auto },
                new TableCellMarginDefault(
                    new TopMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                    new LeftMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                    new RightMargin { Width = "100", Type = TableWidthUnitValues.Dxa }
                )
            ),
            new Paragraph(
                new ParagraphProperties(new Justification { Val = align }),
                new Run(new RunProperties(
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                    new FontSize { Val = "20" }
                ), new Text(text ?? "-") { Space = SpaceProcessingModeValues.Preserve })
            )
        );
    }
    private Paragraph CreateNormalParagraphWithBg(string text, string hexColor)
    {
        var cellColor = new Shading
        {
            Val = ShadingPatternValues.Clear,
            Color = "auto",
            Fill = hexColor.Replace("#", "")
        };

        return new Paragraph(
            new Run(CreateDefaultRunProperties(), new Text(text))
        );
    }
    private TableRow CreateTableRow(params string[] texts)
    {
        var row = new TableRow();
        foreach (var text in texts)
        {
            row.Append(new TableCell(
                new Paragraph(
                    new Run(CreateDefaultRunProperties(), new Text(text))
                )
            ));
        }
        return row;
    }
    private Paragraph CreateEmptyLine()
    {
        return new Paragraph(new Run(new Text(""))); // บรรทัดเปล่า
    }
    private IEnumerable<Paragraph> CreateNumberedList(string[] items)
    {
        int idx = 1;
        foreach (var item in items)
        {
            yield return new Paragraph(
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                        new FontSize { Val = "20" }
                    ),
                    new Text($"{idx++}) {item}")
                )
            );
        }
    }

    private Table CreateHeaderBoxTable(WorkflowPoint workflow)
    {
        var table = new Table(new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single },
                new BottomBorder { Val = BorderValues.Single },
                new LeftBorder { Val = BorderValues.Single },
                new RightBorder { Val = BorderValues.Single },
                new InsideHorizontalBorder { Val = BorderValues.Single },
                new InsideVerticalBorder { Val = BorderValues.Single }
            )));

        var row = new TableRow();
        row.Append(CreateCell("ครั้งที่แก้ไข: " + workflow.EditNumber, JustificationValues.Left));
        row.Append(CreateCell("วันที่แก้ไข: " + workflow.EditDate.ToString("d MMM yy", new CultureInfo("th-TH")), JustificationValues.Left));
        row.Append(CreateCell("หน้า: " + workflow.PageNumber, JustificationValues.Left));
        table.Append(row);
        return table;
    }

    private Table CreateApprovalTable(List<WorkflowApproval> approvals)
    {
        var table = new Table(new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single },
                new BottomBorder { Val = BorderValues.Single },
                new LeftBorder { Val = BorderValues.Single },
                new RightBorder { Val = BorderValues.Single },
                new InsideHorizontalBorder { Val = BorderValues.Single },
                new InsideVerticalBorder { Val = BorderValues.Single }
            )));

        table.Append(CreateApprovalRow("การอนุมัติเอกสาร", "", ""));
        table.Append(CreateApprovalRow("", "ผู้จัดทำ", "ผู้ตรวจสอบ", "ผู้อนุมัติ"));

        var signRow = new TableRow();
        var nameRow = new TableRow();
        var posRow = new TableRow();

        foreach (var approval in approvals.OrderBy(a => a.Level))
        {
            signRow.Append(new TableCell(new Paragraph(new Run(CreateDefaultRunProperties(), new Text(approval.SignText)))));
            nameRow.Append(new TableCell(new Paragraph(new Run(CreateDefaultRunProperties(), new Text($"({approval.FullName})")))));
            posRow.Append(new TableCell(new Paragraph(new Run(CreateDefaultRunProperties(), new Text(approval.Position)))));
        }

        table.Append(signRow);
        table.Append(nameRow);
        table.Append(posRow);
        return table;
    }
    private TableRow CreateApprovalRow(params string[] texts)
    {
        var row = new TableRow();
        foreach (var text in texts)
        {
            var cell = new TableCell(
                new Paragraph(
                    new Run(CreateDefaultRunProperties(), new Text(text))
                )
            );
            cell.TableCellProperties = new TableCellProperties(
                new TableCellMargin
                {
                    TopMargin = new TopMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                    BottomMargin = new BottomMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                    LeftMargin = new LeftMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                    RightMargin = new RightMargin { Width = "100", Type = TableWidthUnitValues.Dxa }
                }
            );
            row.Append(cell);
        }
        return row;
    }
    private Table CreateHistoryTable(List<WorkflowHistory> historyEdits)
    {
        var table = new Table(new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single },
                new BottomBorder { Val = BorderValues.Single },
                new LeftBorder { Val = BorderValues.Single },
                new RightBorder { Val = BorderValues.Single },
                new InsideHorizontalBorder { Val = BorderValues.Single },
                new InsideVerticalBorder { Val = BorderValues.Single }
            )
        ));

        // Add header row
        table.Append(CreateApprovalRow("ครั้งที่แก้ไข", "วันที่แก้ไข", "รายละเอียดการแก้ไข"));

        // Add rows for each history edit
        foreach (var history in historyEdits)
        {
            table.Append(CreateApprovalRow(
                history.EditNumber.ToString(),
                history.EditDate.ToString("d MMM yy", new CultureInfo("th-TH")),
                history.Description
            ));
        }

        return table;
    }
    private void StyleHeader(ExcelRange cell, bool bold = false)
    {
        cell.Style.Font.Bold = bold;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    }
    private Paragraph CreateHeading(string text, int fontSize)
    {
        return new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center }
            ),
            new Run(
                new RunProperties(
                    new Bold(),
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                    new FontSize { Val = fontSize.ToString() }
                ),
                new Text(text)
            )
        );
    }
    private Paragraph CreateHeadingLeft(string text, int fontSize)
    {
        return new Paragraph(
            new ParagraphProperties(
                new Indentation { Left = "500" }
            ),
            new Run(
                new RunProperties(
                    new Bold(),
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                    new FontSize { Val = fontSize.ToString() }
                ),
                new Text(text)
            )
        );
    }

    private Paragraph CreateNumberedParagraph(string title, string[] items)
    {
        var para = new Paragraph(
            new ParagraphProperties(
                new Indentation { Left = "500" }
            ),
            new Run(
                new RunProperties(
                    new Bold(),
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                    new FontSize { Val = "20" }
                ),
                new Text(title)));

        foreach (var item in items.Select((text, i) => $"{i + 1}. {text}"))
        {
            para.Append(
                new Run(new Break()),
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                        new FontSize { Val = "20" }
                    ),
                    new Text(item)
                )
            );
        }
        return para;
    }
    private Table CreateThreeColumnTable(string fiscalYearPrev, string fiscalYear, List<string> prevProcesses, List<string> currentProcesses, List<string> controlActivities)
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

        // Header row
        table.Append(CreateTableRowStyled(
            new[] { $"กระบวนการ ปี {fiscalYearPrev} (เดิม)", $"กระบวนการ ปี {fiscalYear} (ทบทวน)", "กิจกรรมควบคุม (Control Activity)" },
            "DDEBF7"
        ));


        // Data rows
        int rowCount = Math.Max(Math.Max(prevProcesses.Count, currentProcesses.Count), controlActivities.Count);
        for (int i = 0; i < rowCount; i++)
        {
            table.Append(CreateTableRowStyled(new[] {
                i < prevProcesses.Count ? prevProcesses[i] : "",
                i < currentProcesses.Count ? currentProcesses[i] : "",
                i < controlActivities.Count ? controlActivities[i] : "" },
                "FFFFFF"
            ));
        }

        return table;
    }
    private TableRow CreateTableRowStyled(string[] texts, string? bgColorHex = null)
    {
        var row = new TableRow();
        foreach (var text in texts)
        {
            var cellProps = new TableCellProperties();

            if (!string.IsNullOrEmpty(bgColorHex))
            {
                cellProps.Append(new Shading
                {
                    Val = ShadingPatternValues.Clear,
                    Color = "auto",
                    Fill = bgColorHex.Replace("#", "") // remove # if provided
                });
            }

            var cell = new TableCell(
                cellProps,
                new Paragraph(
                    new Run(
                        new RunProperties(
                            new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                            new FontSize { Val = "20" } // 16pt
                        ),
                        new Text(text ?? "")
                    )
                )
            );
            row.Append(cell);
        }
        return row;
    }
    private Paragraph CreateItalicNote(string text)
    {
        return new Paragraph(new ParagraphProperties(
            new Indentation { Left = "500" }
        ),
            new Run(CreateDefaultRunProperties(italic: true), new Text(text))
        );
    }
    private Paragraph CreateCheckBoxOptions(string[] options)
    {
        var para = new Paragraph();
        foreach (var opt in options)
        {
            para.Append(
                new ParagraphProperties(new Indentation { Left = "720" }),
                new Run(CreateDefaultRunProperties(), new Text("☐ " + opt)),
                new Run(new Break())
            );
        }
        return para;
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

    private Table CreateSignatureSection(
        string leftName, string leftPosition,
        string rightName, string rightPosition,
        string leftDate, string rightDate,
        TableRowAlignmentValues alignment
    )
    {
        var table = new Table(
            new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }, // Adjusted width
                new TableJustification { Val = alignment } // Use the passed alignment value
            ),
            new Run(
                new RunProperties(
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                    new FontSize { Val = "20" }
                )
            )
        );

        var row1 = new TableRow();
        row1.Append(
            CreateCell(leftName),
            CreateCell(rightName)
        );

        var row2 = new TableRow();
        row2.Append(
            CreateCell(leftPosition),
            CreateCell(rightPosition)
        );

        var row3 = new TableRow();
        row3.Append(
            CreateCell("วันที่ " + leftDate),
            CreateCell("วันที่ " + rightDate)
        );

        table.Append(row1, row2, row3);
        return table;
    }

    private TableCell CreateCell(string text)
    {
        return new TableCell(
            new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Center }
                ),
                new Run(
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                    new FontSize { Val = "20" },
                    new Text(text ?? "")
                )
            )
        );
    }
    private Paragraph CreateBoldParagraph(string text, int fontSize)
    {
        return new Paragraph(new ParagraphProperties(new Indentation { Left = "500" }),
            new Run(
                new RunProperties(
                    new Bold(),
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                    new FontSize { Val = fontSize.ToString() }
                ),
                new Text(text)
            )
        );
    }
    private Paragraph CreateNormalParagraph(string text)
    {
        return new Paragraph(new ParagraphProperties(
            new Indentation { Left = "720" }
        ), new Run(
            new RunProperties(
                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
                new FontSize { Val = "20" }
            ),
            new Text(text ?? "")
        ));
    }
    private Paragraph CreatePageBreak()
    {
        return new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
    }
    private RunProperties CreateDefaultRunProperties(bool bold = false, bool italic = false)
    {
        var runProps = new RunProperties(
            new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK" },
            new FontSize { Val = "20" }
        );
        if (bold) runProps.Append(new Bold());
        if (italic) runProps.Append(new Italic());
        return runProps;
    }


}
