using BatchAndReport.Models;
using BatchAndReport.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using System.IO;

public class WordService : IWordService
{
    public byte[] GenerateWord(SMEProjectDetailModels model)
    {
        using var stream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            // หัวเรื่อง
            body.Append(CreateHeading($"แบบฟอร์มการจัดทำข้อเสนอ ประจำปีงบประมาณ {model.FiscalYear}", 20));
            body.Append(CreateEmptyLine());

            body.Append(CreateNormalParagraph($"กระทรวง : {model.MinistryName}"));
            body.Append(CreateNormalParagraph($"หน่วยงาน : {model.DepartmentName}"));
            body.Append(CreateNormalParagraph($"ชื่อกิจกรรม : {model.ActivityName}"));
            body.Append(CreateNormalParagraph($"งบประมาณ : {model.BudgetAmount:N0}"));

            body.Append(CreateEmptyLine());

            body.Append(CreateBoldParagraph("□ ใช้งบประมาณ"));
            //foreach (var plan in model.Plans)
                body.Append(CreateIndentedParagraph("   □ " + "xxxxxxxxxxx"));
                body.Append(CreateIndentedParagraph("   □ " + "xxxxxxxxxxx"));
            body.Append(CreateNormalParagraph("□ ไม่ใช้งบประมาณ"));

            body.Append(CreateBoldParagraph("สถานภาพโครงการ : □ โครงการใหม่ □ โครงการต่อเนื่อง □ โครงการเดิม □ โครงการ Flagship"));
            body.Append(CreateEmptyLine());

            var table = new Table(new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }
            ,   new TableBorders(
                new TopBorder { Val = BorderValues.Single },
                new BottomBorder { Val = BorderValues.Single },
                new LeftBorder { Val = BorderValues.Single },
                new RightBorder { Val = BorderValues.Single },
                new InsideHorizontalBorder { Val = BorderValues.Single },
                new InsideVerticalBorder { Val = BorderValues.Single })));
            table.Append(CreateTableRowColor("", "ผู้รับผิดชอบโครงการ", "ผู้ประสานงาน"));
            table.Append(CreateTableRow("ชื่อ-นามสกุล", model.OwnerName, model.ContactName));
            table.Append(CreateTableRow("ตำแหน่ง", model.OwnerPosition, model.ContactPosition));
            table.Append(CreateTableRow("โทรศัพท์", model.OwnerPhone, model.ContactPhone));
            table.Append(CreateTableRow("มือถือ", model.OwnerMobile, model.ContactMobile));
            table.Append(CreateTableRow("Email", model.OwnerEmail, model.ContactEmail));
            table.Append(CreateTableRow("Line ID", model.OwnerLineId, model.ContactLineId));
            body.Append(table);

            body.Append(CreateBoldParagraph("ประเด็นสำคัญในการส่งเสริม SME ปี พ.ศ.{model.FiscalYear}"));
            //foreach (var item in model.PromotionStrategies)
            //body.Append(CreateNormalParagraph("□ " + item));
            body.Append(CreateBoldParagraph("□ Digital □ Environment/Green □ Social □ Governance □ Soft power"));

            body.Append(CreateNormalParagraph($"ประเด็นสำคัญในการส่งเสริม SME ปี พ.ศ.{model.FiscalYear} ประเด็นการส่งเสริม/กลยุทธ์ที่สอดคล้องกับแผนปฎิบัติการส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อมประจำปีงบประมาณ (เลือกเพียง 1 ประเด็นการส่งเสริม 1 กลยุทธ์ ต่อ 1 โครงการ)"));
            // Header
            table.Append(CreateTableRowColor("ประเด็นการส่งเสริม", "กลยุทธ์"));

            // Rows
            foreach (var group in model.Strategies)
            {
                var first = true;
                int index = 1;

                foreach (var item in model.Strategies)
                {
                    var row = new TableRow();
                    if (first)
                    {
                        row.Append(new TableCell(
                            new TableCellProperties(new VerticalMerge { Val = MergedCellValues.Restart }),
                            new Paragraph(new Run(new Text($"□ {group.StrategyId}")))));
                        first = false;
                    }
                    else
                    {
                        row.Append(new TableCell(
                            new TableCellProperties(new VerticalMerge { Val = MergedCellValues.Continue }),
                            new Paragraph(new Run()))); // Empty cell
                    }

                    row.Append(new TableCell(new Paragraph(new Run(new Text($"□ {index++} {item.StrategyDesc}")))));
                    table.Append(row);
                }
            }

            body.Append(table);
            body.Append(CreateBoldParagraph("ความสำคัญของโครงการ/หลักการและเหตุผล :"));
            body.Append(CreateNormalParagraph(model.ProjectRationale ?? ""));

            body.Append(CreateBoldParagraph("วัตถุประสงค์ของโครงการ :"));
            body.Append(CreateNormalParagraph(model.ProjectObjective ?? ""));

            body.Append(CreateBoldParagraph("กลุ่มเป้าหมาย (สามารถเลือกได้มากกว่า 1 กลุ่มเป้าหมาย):"));
            //foreach (var group in model.TargetGroups)
                //body.Append(CreateNormalParagraph("□ " + group));
                body.Append(CreateIndentedParagraph("□ วิสาหกิจระยะเริ่มต้น Early-Stage Enterprise □ วิสาหกิจขนาดย่อม Small Enterprise"));
                body.Append(CreateIndentedParagraph("□ วิสาหกิจรายย่อย Micro Enterprise □ วิสาหกิจขนาดกลาง Medium Enterprise □ ทุกกลุ่ม"));

            body.Append(CreateBoldParagraph("รายละเอียดแผนการดำเนินงาน/กิจกรรม (โปรดอธิบายขั้นตอนการดำเนินงานในแต่ละกิจกรรมของโครงงานทั้งหมด จำแนกเป็นข้อๆ ตามลำดับขั้นตอนการไหลของงาน โดยละเอียด)\r\n :"));
            body.Append(CreateNormalParagraph(model.Activities ?? ""));

            body.Append(CreateBoldParagraph("จุดเด่นของโครงการ (อธิบายภาพรวมโดยย่อ และแสดงให้เห็นถึงจุดเด่นและความสำคัญของโครงการ) :"));
            body.Append(CreateNormalParagraph(model.ProjectFocus ?? ""));

            body.Append(CreateBoldParagraph("พื้นที่ดำเนินการ :(ระบุภาค/พื้นที่เป้าหมาย)"));
                if (model.OperationArea != null && model.OperationArea.Any())
                {
                    body.Append(CreateNormalParagraph(string.Join(", ", model.OperationArea)));
                }
                else
                {
                    body.Append(CreateNormalParagraph(""));
                }

            body.Append(CreateBoldParagraph("สาขาเป้าหมาย :(ระบุสาขาเป้าหมาย เช่น ทุกสาขา สาขาท่องเที่ยว สาขาอาหารแปรรูป สาขายานยนต์และชิ้นส่วน สาขาอัญมณีและเครื่องประดับ เป็นต้น)"));
                if (model.IndustrySector != null && model.IndustrySector.Any())
                {
                    body.Append(CreateNormalParagraph(string.Join(", ", model.IndustrySector)));
                }
                else
                {
                    body.Append(CreateNormalParagraph(""));
                }

            body.Append(CreateBoldParagraph("การพัฒนา 11 อุตสาหกรรม Soft Power :(ระบุสาขาอุตสาหกรรม หากท่านเลือกกลยุทธ์ที่ 16)"));
            //foreach (var power in model.SoftPowers)
            //    body.Append(CreateNormalParagraph("□ " + power));
                  body.Append(CreateIndentedParagraph("□ " + "power 1"));
                  body.Append(CreateIndentedParagraph("□ " + "power 2"));
                  body.Append(CreateIndentedParagraph("□ " + "power 3"));

            body.Append(CreateBoldParagraph("ระยะเวลาในการดำเนินโครงการ :"));
            body.Append(CreateNormalParagraph(model.Timeline ?? ""));
            body.Append(CreateBoldParagraph("หน่วยงานที่ร่วมบูรณาการ รูปแบบ/วิธีร่วมดำเนินการร่วมกัน (รูปแบบการดำเนินงานแบบบูรณาการ การส่งต่อผู้ประกอบการ ฯลฯ โปรดระบุ) :"));
            body.Append(CreateNormalParagraph(model.OrgPartner ?? "" + "ทำหน้าที่" + model.RoleDescription));

            body.Append(CreateBoldParagraph("ตัวชี้วัดที่สำคัญ (โปรดระบุตัวชี้วัดระดับผลผลิต และผลลัพธ์ของโครงการ พร้อมหน่วยนับแลัเป้าหมายเชิงปริมาณ)"));
            var mainMetricTable = new Table(new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
                new TableBorders(
                new TopBorder { Val = BorderValues.Single },
                new BottomBorder { Val = BorderValues.Single },
                new LeftBorder { Val = BorderValues.Single },
                new RightBorder { Val = BorderValues.Single },
                new InsideHorizontalBorder { Val = BorderValues.Single },
                new InsideVerticalBorder { Val = BorderValues.Single })));
            mainMetricTable.Append(CreateTableRowColor("ตัวชี้วัดผลผลิต", "จำนวนเป้าหมาย", "หน่วยนับ", "วิธีการวัดผล"));
            foreach (var item in model.OutputIndicators)
                mainMetricTable.Append(CreateTableRow(item.Name, item.Target, item.Unit, item.Method));
            //mainMetricTable.Append(CreateTableRow("item.Name", "item.Target", "item.Unit", "item.Method"));
            //mainMetricTable.Append(CreateTableRow("item.Name", "item.Target", "item.Unit", "item.Method"));
            //mainMetricTable.Append(CreateTableRow("item.Name", "item.Target", "item.Unit", "item.Method"));
            body.Append(mainMetricTable);

            body.Append(CreateEmptyLine());
            var outcomeMetricTable = new Table(new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }, 
                new TableBorders(
                new TopBorder { Val = BorderValues.Single },
                new BottomBorder { Val = BorderValues.Single },
                new LeftBorder { Val = BorderValues.Single },
                new RightBorder { Val = BorderValues.Single },
                new InsideHorizontalBorder { Val = BorderValues.Single },
                new InsideVerticalBorder { Val = BorderValues.Single })));
            outcomeMetricTable.Append(CreateTableRowColor("ตัวชี้วัดผลลัพธ์", "จำนวนเป้าหมาย", "หน่วยนับ", "วิธีการวัดผล"));
            foreach (var item in model.OutcomeIndicators)
                outcomeMetricTable.Append(CreateTableRow(item.Name, item.Target, item.Unit, item.Method));
            //outcomeMetricTable.Append(CreateTableRow("item.Name", "item.Target", "item.Unit", "item.Method"));
            //outcomeMetricTable.Append(CreateTableRow("item.Name", "item.Target", "item.Unit", "item.Method"));
            //outcomeMetricTable.Append(CreateTableRow("item.Name", "item.Target", "item.Unit", "item.Method"));
            body.Append(outcomeMetricTable);

            body.Append(CreateNormalParagraph("ข้อมูลอื่นๆ เพิ่มเติมที่จะช่วยสร้างความเข้าใจเกี่ยวกับการดำเนินงาน :"));
            body.Append(CreateNormalParagraph(model.AdditionalNotes ?? ""));

            mainPart.Document.Save();
        }

        return stream.ToArray();
    }

    public byte[] GenerateSummaryWord(
    List<SMESummaryProjectModels> projects,
    List<SMEStrategyDetailModels> strategyList,
    string year)
    {
        var culture = new CultureInfo("th-TH");
        int totalProjects = projects.Sum(p => p.ProjectCount ?? 0);
        decimal totalBudget = projects.Sum(p => p.Budget ?? 0);

        using var stream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            // ------------------ Part 1: Summary ------------------
            body.Append(CreateBoldParagraphAlign(
                $"ภาพรวมโครงการและงบประมาณเพื่อการส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (SME)\n" +
                $"ภายใต้แผนปฏิบัติการส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ประจำปี พ.ศ. {year}",
                18, JustificationValues.Center));

            body.Append(CreateHorizontalLine());

            var table = new Table(new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single },
                    new BottomBorder { Val = BorderValues.Single },
                    new LeftBorder { Val = BorderValues.Single },
                    new RightBorder { Val = BorderValues.Single },
                    new InsideHorizontalBorder { Val = BorderValues.Single },
                    new InsideVerticalBorder { Val = BorderValues.Single })));

            table.Append(CreateHeaderRow("ประเด็นการส่งเสริม SME", "จำนวนโครงการ", "งบประมาณ (ล้านบาท)"));

            int i = 1;
            foreach (var row in projects)
            {
                table.Append(CreateDataRow(false,
                    $"ประเด็นที่{i} {row.IssueName ?? ""}",
                    row.ProjectCount?.ToString("N0", culture) ?? "0",
                    (row.Budget.GetValueOrDefault() / 1_000_000).ToString("N4", culture)));
                i++;
            }

            table.Append(CreateDataRow(true,
                "รวมทั้งหมด",
                totalProjects.ToString("N0", culture),
                (totalBudget / 1_000_000).ToString("N4", culture)));

            body.Append(table);
            body.Append(CreateEmptyParagraph());
            body.Append(CreateNormalParagraph("โดยมีหน่วยงานทั้งหมด xx กระทรวง xx หน่วยงาน"));

            // ------------------ Part 2: Strategy Detail ------------------
            var grouped = strategyList
            .GroupBy(p => p.Topic)
            .ToList();

            int topicIndex = 1;
            foreach (var topicGroup in grouped)
            {
                body.Append(CreateBoldParagraph($"ประเด็นการส่งเสริมที่ {topicIndex} {topicGroup.Key}"));

                var strategyGrouped = topicGroup.GroupBy(p => p.StrategyDesc).ToList();
                int strategyIndex = 1;
                foreach (var strategyGroup in strategyGrouped)
                {
                    var totalProject = strategyGroup.Count();
                    var sumBudget = strategyGroup.Sum(p => p.BudgetAmount);

                    body.Append(CreateBoldParagraph($"กลยุทธ์ที่ {strategyIndex} {strategyGroup.Key}"));
                    body.Append(CreateNormalParagraph($"จำนวน {totalProject} โครงการ งบประมาณ {sumBudget:N2} ล้านบาท"));

                    // Table
                    var tableStrategy = new Table(new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single },
                    new BottomBorder { Val = BorderValues.Single },
                    new LeftBorder { Val = BorderValues.Single },
                    new RightBorder { Val = BorderValues.Single },
                    new InsideHorizontalBorder { Val = BorderValues.Single },
                    new InsideVerticalBorder { Val = BorderValues.Single })));
                    tableStrategy.Append(CreateHeaderRow("หน่วยงาน/โครงการ", "งบประมาณ"));

                    // Fix for CS1061: Replace 'NameTh' with 'Department' in the GroupBy clause
                    var deptGrouped = strategyGroup
                        .GroupBy(p => new { p.DepartmentCode, p.Department })
                        .ToList();

                    int projectIndex = 1;
                    foreach (var deptGroup in deptGrouped)
                    {
                        var deptTotal = deptGroup.Sum(p => p.BudgetAmount);
                        tableStrategy.Append(CreateDataRowColor(true, deptGroup.Key.Department, $"{deptTotal:N2}"));

                        foreach (var proj in deptGroup)
                        {
                            tableStrategy.Append(CreateDataRow(false, $"   {projectIndex}. {proj.ProjectName}", $"{proj.BudgetAmount:N2}"));
                            projectIndex++;
                        }
                    }

                    body.Append(tableStrategy);
                    body.Append(CreateEmptyParagraph());
                    strategyIndex++;
                }

                topicIndex++;
                body.Append(CreatePageBreak());
            }

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

    private Paragraph CreateHeading(string text, int fontSize)
    {
        return new Paragraph(
            new Run(
                new RunProperties(
                    new Bold(),
                    new FontSize { Val = (fontSize * 2).ToString() }  // Word uses half-point units
                ),
                new Text(text)
            )
        );
    }
    private Paragraph CreatePageBreak()
    {
        return new Paragraph(
            new Run(
                new Break() { Type = BreakValues.Page }
            )
        );
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

    private TableRow CreateTableRow(params string[] columns)
    {
        var row = new TableRow();
        foreach (var col in columns)
        {
            var cell = new TableCell(new Paragraph(new Run(new Text(col ?? ""))));
            row.Append(cell);
        }
        return row;
    }

    private TableRow CreateTableRowColor(params string[] columns)
    {
        var row = new TableRow();
        foreach (var col in columns)
        {
            var cell = new TableCell(
                new TableCellProperties(
                    new Shading
                    {
                        Val = ShadingPatternValues.Clear,
                        Color = "auto",
                        Fill = "BDD7EE" // Light blue background
                    }
                ),
                new Paragraph(new Run(new Text(col ?? "")))
            );
            row.Append(cell);
        }
        return row;
    }

    private Paragraph CreateNormalParagraph(string text)
    {
        return new Paragraph(new Run(new Text(text ?? "")));
    }

    private Paragraph CreateEmptyLine()
    {
        return new Paragraph(new Run(new Text(" ")));
    }
    private Paragraph CreateIndentedParagraph(string text, int leftChars = 1)
    {
        return new Paragraph(
            new ParagraphProperties(
                new Indentation() { Left = (leftChars * 360).ToString() } // 360 twips ≈ 0.25 inch per char
            ),
            new Run(new Text(text ?? ""))
        );
    }

    private TableRow CreateHeaderRow(params string[] headers)
    {
        var row = new TableRow();
        foreach (var header in headers)
        {
            var cellProps = new TableCellProperties(
                new Shading
                {
                    Val = ShadingPatternValues.Clear,
                    Color = "auto",
                    Fill = "DAE8FC" // สีพื้นหลัง (Hex) เช่น DAE8FC = ฟ้าอ่อน
                });

            var paragraph = new Paragraph(
                new Run(
                    new RunProperties(new Bold()),
                    new Text(header)
                )
            );

            var cell = new TableCell(cellProps, paragraph);
            row.Append(cell);
        }
        return row;
    }

    private TableRow CreateDataRow(bool isBold, params string[] values)
    {
        var row = new TableRow();
        foreach (var value in values)
        {
            Run run;
            if (isBold)
            {
                run = new Run(new RunProperties(new Bold()), new Text(value ?? ""));
            }
            else
            {
                run = new Run(new Text(value ?? ""));
            }

            row.Append(new TableCell(new Paragraph(run)));
        }
        return row;
    }

    private TableRow CreateDataRowColor(bool isBold, params string[] values)
    {
        var row = new TableRow();
        foreach (var value in values)
        {
            var run = isBold
                ? new Run(new RunProperties(new Bold()), new Text(value ?? ""))
                : new Run(new Text(value ?? ""));

            var paragraph = new Paragraph(run);

            var cellProperties = new TableCellProperties();
            if (isBold)
            {
                // สีเหลืองอ่อน (Hex: FFFF99) หรือใช้ Hex สีอื่นได้
                cellProperties.Append(new Shading
                {
                    Color = "auto",
                    Fill = "FFFF99",   // สีพื้นหลัง
                    Val = ShadingPatternValues.Clear
                });
            }

            var cell = new TableCell();
            cell.Append(cellProperties);
            cell.Append(paragraph);
            row.Append(cell);
        }
        return row;
    }

    private Paragraph CreateBoldParagraphAlign(string text, int fontSize = 22, JustificationValues align = default)
    {
        var justification = align == default ? new Justification { Val = JustificationValues.Left } : new Justification { Val = align };

        return new Paragraph(
            new ParagraphProperties(justification),
            new Run(
                new RunProperties(
                    new Bold(),
                    new FontSize { Val = (fontSize * 2).ToString() }),
                new Text(text) { Space = SpaceProcessingModeValues.Preserve })
        );
    }
    private Paragraph CreateEmptyParagraph()
    {
        return new Paragraph(new Run(new Text(" ")));
    }
    private Paragraph CreateHorizontalLine()
    {
        return new Paragraph(new Run(new Text(new string('*', 70))));
    }
    private TableRow CreateDataRow(string left, string right)
    {
        return new TableRow(
            new TableCell(new Paragraph(new Run(new Text(left)))),
            new TableCell(new Paragraph(new Run(new Text(right))))
        );
    }
}
