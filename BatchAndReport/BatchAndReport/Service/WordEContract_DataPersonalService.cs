using BatchAndReport.DAO;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Http;
using System.Text;

public class WordEContract_DataPersonalService
{
    private readonly WordServiceSetting _w;
    private readonly Econtract_Report_PDSADAO _e;
    private readonly IConverter _pdfConverter;
    private readonly E_ContractReportDAO _eContractReportDAO;
    public WordEContract_DataPersonalService(WordServiceSetting ws, Econtract_Report_PDSADAO e
        , IConverter pdfConverter
        , E_ContractReportDAO eContractReportDAO
        )
    {
        _w = ws;
        _e = e;
        _pdfConverter = pdfConverter;
        _eContractReportDAO = eContractReportDAO;
    }

    #region  4.1.1.2.6.บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล PDSA
    public async Task<byte[]> OnGetWordContact_DataPersonalService(string id)
    {
        try
        {

            var result = await _e.GetPDSAAsync(id);
            var rLe = await _e.GetPDSA_LegalBasisSharingAsync(id);
            var rSd = await _e.GetPDSA_Shared_DataAsync(id);
            var stream = new MemoryStream();

            using (var wordDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();

                // Styles
                var stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylePart.Styles = WordServiceSetting.CreateDefaultStyles();

                var body = mainPart.Document.AppendChild(new Body());

                // 1. Logo (centered)
                var imagePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
                var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg, "rIdLogo");
                using (var imgStream = File.OpenRead(imagePath))
                {
                    imagePart.FeedData(imgStream);
                }

                // --- 1. Top Row: Logo left, Contract code box right ---
                var topTable = new Table(
                    new TableProperties(
                        new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                        new TableBorders(
                            new TopBorder { Val = BorderValues.None },
                            new BottomBorder { Val = BorderValues.None },
                            new LeftBorder { Val = BorderValues.None },
                            new RightBorder { Val = BorderValues.None },
                            new InsideHorizontalBorder { Val = BorderValues.None },
                            new InsideVerticalBorder { Val = BorderValues.None }
                        )
                    ),
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(
                                new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "60" }
                            ),
                            new Paragraph(
                                new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
                                WordServiceSetting.CreateImage(
                                    mainPart.GetIdOfPart(imagePart),
                                    240, 80
                                )
                            )
                        ),
                        new TableCell(
                            new TableCellProperties(
                                new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "40" }
                            ),
                            new Paragraph(
                                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                                WordServiceSetting.CreateImage(
                                    mainPart.GetIdOfPart(imagePart),
                                    240, 80
                                )
                            )
                        )
                    )
                );
                body.AppendChild(topTable);

                // --- 2. Titles ---
                body.AppendChild(WordServiceSetting.EmptyParagraph());
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("ข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล", "32"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("(Personal Data Sharing Agreement)", "32"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("ระหว่าง", "32"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม กับ " + result.ContractPartyName + " ", "32"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("---------------------------------------------", "32"));

                string datestring = CommonDAO.ToThaiDateStringCovert(result.Master_Contract_Sign_Date ?? DateTime.Now);
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล (“ข้อตกลง”) ฉบับนี้ทำขึ้น เมื่อ " + datestring + " ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("โดยที่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สสว.” ฝ่ายหนึ่ง ได้ตกลงใน.... (ระบุชื่อบันทึกข้อตกลงความร่วมมือ/สัญญาหลัก) .... สัญญาเลขที่ .......... (ระบุเลขที่บันทึกข้อตกลงความร่วมมือ/สัญญาหลัก).................  ฉบับลงวันที่ ..... (ระบุวันที่ลงนามข้อตกลงความร่วมมือหรือวันทำสัญญาหลัก) .......... ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สัญญาหลัก” กับ ........ (ระบุชื่อคู่สัญญาอีกฝ่าย) ........ ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “..... (ระบุชื่อเรียกคู่สัญญาอีกฝ่าย ......” อีกฝ่ายหนึ่ง รวมเรียกว่า “คู่สัญญา”", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("เพื่อให้บรรลุวัตถุประสงค์ภายใต้ความตกลงของสัญญาหลัก คู่สัญญามีความจำเป็นต้องแบ่งปัน โอน แลกเปลี่ยน หรือเปิดเผย (รวมเรียกว่า “แบ่งปัน”) ข้อมูลส่วนบุคคลที่ตนเก็บรักษาแก่อีกฝ่าย ซึ่งข้อมูลส่วนบุคคลที่แต่ละฝ่าย เก็บรวมรวม ใช้หรือเปิดเผย (รวมเรียกว่า “ประมวลผล”) นั้น แต่ละฝ่ายต่างเป็นผู้ควบคุมข้อมูลส่วนบุคคล ตามกฎหมายที่เกี่ยวข้องกับการคุ้มครองข้อมูลส่วนบุคคล กล่าวคือแต่ละฝ่ายต่างเป็นผู้มีอำนาจตัดสินใจ กำหนดรูปแบบ และกำหนดวัตถุประสงค์ ในการประมวลผลข้อมูลส่วนบุคคลในข้อมูลที่ตนต้องแบ่งปัน ภายใต้ข้อตกลงนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ด้วยเหตุนี้ คู่สัญญาจึงตกลงจัดทำข้อตกลงฉบับนี้ และให้ถือเป็นส่วนหนึ่งของสัญญาหลัก เพื่อเป็นหลักฐานการแบ่งปันข้อมูลส่วนบุคคลระหว่างคู่สัญญาและเพื่อดำเนินการให้เป็นไปตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒ และกฎหมายอื่น ๆ ที่ออกตามความในพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒ ซึ่งต่อไปในข้อตกลงฉบับนี้ รวมเรียกว่า “กฎหมายคุ้มครองข้อมูลส่วนบุคคล”  ทั้งที่มีผลใช้บังคับอยู่ ณ วันทำข้อตกลงฉบับนี้ และที่จะมีการเพิ่มเติมหรือแก้ไขเปลี่ยนแปลงในภายหลัง โดยมีรายละเอียดดังนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("1.คู่สัญญารับทราบว่า ข้อมูลส่วนบุคคล หมายถึง ข้อมูลเกี่ยวกับบุคคลธรรมดา ซึ่งทำให้สามารถระบุตัวบุคคลนั้นได้ไม่ว่าทางตรงหรือทางอ้อม โดยคู่สัญญาแต่ละฝ่าย จะดำเนินการตามที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลกำหนด เพื่อคุ้มครองให้การประมวลผลข้อมูลส่วนบุคคลเป็นไปอย่างเหมาะสมและถูกต้องตามกฎหมาย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("2.ข้อมูลส่วนบุคคลที่คู่สัญญาแบ่งปันกัน คู่สัญญาแต่ละฝ่ายตกลงแบ่งปันข้อมูลส่วนบุคคลดังรายการต่อไปนี้แก่คู่สัญญาอีกฝ่าย", null, "32"));

                // Table: ข้อมูลส่วนบุคคลที่แบ่งปันโดย สสว.
                var rleData = rSd.Where(e => e.Owner == "OSMEP").ToList();
                if (rleData != null && rleData.Count() > 0)
                {
                    var infoTable = new Table(
                        new TableProperties(
                               new TableWidth { Width = "10000", Type = TableWidthUnitValues.Dxa }, // ขยายตาราง
                            new TableBorders(
                                new TopBorder { Val = BorderValues.Single, Size = 4 },
                                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                                new RightBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                            )
                        ),
                        // Header row
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new Shading { Fill = "D9D9D9" }),
                                WordServiceSetting.BoldParagraph("ข้อมูลส่วนบุคคลที่แบ่งปันโดย สสว.")
                            ),
                            new TableCell(
                                new TableCellProperties(new Shading { Fill = "D9D9D9" }),
                                WordServiceSetting.BoldParagraph("วัตถุประสงค์ในการแบ่งปันข้อมูลส่วนบุคคล")
                            )
                        )
                    );

                    foreach (var item in rleData)
                    {
                        infoTable.AppendChild(
                            new TableRow(
                                new TableCell(
                                    WordServiceSetting.NormalParagraph(item.Detail ?? "-", null, "32")
                                ),
                                new TableCell(
                                    WordServiceSetting.NormalParagraph(item.Objective ?? "-", null, "32")
                                )
                            )
                        );
                    }

                    body.AppendChild(infoTable);
                }
                else
                {
                    var infoTable = new Table(
                        new TableProperties(
                            new TableBorders(
                                new TopBorder { Val = BorderValues.Single, Size = 4 },
                                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                                new RightBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                            )
                        ),
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new Shading { Fill = "D9D9D9" }),
                                WordServiceSetting.BoldParagraph("ข้อมูลส่วนบุคคลที่แบ่งปันโดย สสว.")
                            ),
                            new TableCell(
                                new TableCellProperties(new Shading { Fill = "D9D9D9" }),
                                WordServiceSetting.BoldParagraph("วัตถุประสงค์ในการแบ่งปันข้อมูลส่วนบุคคล")
                            )
                        ),
                        new TableRow(
                            new TableCell(
                                WordServiceSetting.NormalParagraph("-", null, "32")
                            ),
                            new TableCell(
                                WordServiceSetting.NormalParagraph("-", null, "32")
                            )
                        )
                    );
                    body.AppendChild(infoTable);
                }

                var rSdData = rSd.Where(e => e.Owner == "CP").ToList();
                if (rSdData != null && rSdData.Count() > 0)
                {
                    var infoTable = new Table(
                           new TableWidth { Width = "10000", Type = TableWidthUnitValues.Dxa }, // ขยายตาราง
                        new TableProperties(
                            new TableBorders(
                                new TopBorder { Val = BorderValues.Single, Size = 4 },
                                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                                new RightBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                            )
                        ),
                        // Header row
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new Shading { Fill = "D9D9D9" }),
                                WordServiceSetting.BoldParagraph("ข้อมูลส่วนบุคคลที่แบ่งปันโดย สสว.")
                            ),
                            new TableCell(
                                new TableCellProperties(new Shading { Fill = "D9D9D9" }),
                                WordServiceSetting.BoldParagraph("วัตถุประสงค์ในการแบ่งปันข้อมูลส่วนบุคคล")
                            )
                        )
                    );

                    foreach (var item in rSdData)
                    {
                        infoTable.AppendChild(
                            new TableRow(
                                new TableCell(
                                    WordServiceSetting.NormalParagraph(item.Detail ?? "-", null, "32")
                                ),
                                new TableCell(
                                    WordServiceSetting.NormalParagraph(item.Objective ?? "-", null, "32")
                                )
                            )
                        );
                    }

                    body.AppendChild(infoTable);
                }
                else
                {
                    var infoTable = new Table(
                        new TableProperties(
                            new TableBorders(
                                new TopBorder { Val = BorderValues.Single, Size = 4 },
                                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                                new RightBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                            )
                        ),
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new Shading { Fill = "D9D9D9" }),
                                WordServiceSetting.BoldParagraph("ข้อมูลส่วนบุคคลที่แบ่งปันโดย(" + result.ContractPartyName + ")")
                            ),
                            new TableCell(
                                new TableCellProperties(new Shading { Fill = "D9D9D9" }),
                                WordServiceSetting.BoldParagraph("วัตถุประสงค์ในการแบ่งปันข้อมูลส่วนบุคคล")
                            )
                        ),
                        new TableRow(
                            new TableCell(
                                WordServiceSetting.NormalParagraph("-", null, "32")
                            ),
                            new TableCell(
                                WordServiceSetting.NormalParagraph("-", null, "32")
                            )
                        )
                    );
                    body.AppendChild(infoTable);
                }


                // Table: ฐานกฎหมายในการแบ่งปันข้อมูลส่วนบุคคล
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3.ฐานกฎหมายในการแบ่งปันข้อมูลส่วนบุคคล ภายใต้วัตถุประสงค์ที่ระบุในข้อ 2 คู่สัญญาแต่ละฝ่ายมีฐานกฎหมายตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลดังต่อไปนี้ ในการแบ่งปันข้อมูลส่วนบุคคลแก่คู่สัญญาอีกฝ่าย (แต่ละฝ่ายอาจใช้ฐานกฎหมายที่ต่างกันในการแบ่งปันข้อมูลส่วนบุคคล)", null, "32"));

                var OsmepLeg = rLe.Where(e => e.Owner == "OSMEP").ToList();
                if (OsmepLeg != null && OsmepLeg.Count > 0)
                {
                    var legalBasisTable = new Table(
                        new TableProperties(
                                  new TableWidth { Width = "10000", Type = TableWidthUnitValues.Dxa }, // ขยายตาราง
        new TableJustification { Val = TableRowAlignmentValues.Left },      // จัดตรงกลาง
                            new TableBorders(
                                new TopBorder { Val = BorderValues.Single, Size = 4 },
                                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                                new RightBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                            )
                        ),
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new Shading { Fill = "FFFFFF" }),
                                WordServiceSetting.BoldParagraph("ฐานกฎหมายของ สสว.")
                            )
                        )
                    );

                    foreach (var item in OsmepLeg)
                    {
                        legalBasisTable.AppendChild(
                            new TableRow(
                                new TableCell(
                                    WordServiceSetting.NormalParagraph(item.Detail ?? "-", null, "32")
                                )
                            )
                        );
                    }

                    body.AppendChild(legalBasisTable);
                }
                else
                {
                    var legalBasisTable = new Table(
                        new TableProperties(
                                  new TableWidth { Width = "1000", Type = TableWidthUnitValues.Dxa }, // ขยายตาราง
        new TableJustification { Val = TableRowAlignmentValues.Left },      // จัดตรงกลาง
                            new TableBorders(
                                new TopBorder { Val = BorderValues.Single, Size = 4 },
                                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                                new RightBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                            )
                        ),
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new Shading { Fill = "FFFFFF" }),
                                WordServiceSetting.BoldParagraph("ฐานกฎหมายของ สสว.")
                            )
                        ),
                        new TableRow(
                            new TableCell(
                                WordServiceSetting.BoldParagraph("-")
                            )
                        ),
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new Shading { Fill = "FFFFFF" }),
                                WordServiceSetting.BoldParagraph("-")
                            )
                        )
                    );
                    body.AppendChild(legalBasisTable);
                }

                var CPLeg = rLe.Where(e => e.Owner == "OSMEP").ToList();
                if (CPLeg != null && CPLeg.Count > 0)
                {
                    var legalBasisTable = new Table(
                        new TableProperties(
                                  new TableWidth { Width = "10000", Type = TableWidthUnitValues.Dxa }, // ขยายตาราง
        new TableJustification { Val = TableRowAlignmentValues.Left },      // จัดตรงกลาง
                            new TableBorders(
                                new TopBorder { Val = BorderValues.Single, Size = 4 },
                                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                                new RightBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                            )
                        ),
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new Shading { Fill = "FFFFFF" }),
                                WordServiceSetting.BoldParagraph("ฐานกฎหมายของ " + result.ContractPartyName + "")
                            )
                        )
                    );

                    foreach (var item in CPLeg)
                    {
                        legalBasisTable.AppendChild(
                            new TableRow(
                                new TableCell(
                                    WordServiceSetting.NormalParagraph(item.Detail ?? "-", null, "32")
                                )
                            )
                        );
                    }

                    body.AppendChild(legalBasisTable);
                }
                else
                {
                    var legalBasisTable = new Table(
                        new TableProperties(
                                  new TableWidth { Width = "1000", Type = TableWidthUnitValues.Dxa }, // ขยายตาราง
        new TableJustification { Val = TableRowAlignmentValues.Left },      // จัดตรงกลาง
                            new TableBorders(
                                new TopBorder { Val = BorderValues.Single, Size = 4 },
                                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                                new RightBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                            )
                        ),
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new Shading { Fill = "FFFFFF" }),
                                WordServiceSetting.BoldParagraph("ฐานกฎหมายของ " + result.ContractPartyName + "")
                            )
                        ),
                        new TableRow(
                            new TableCell(
                                WordServiceSetting.BoldParagraph("-")
                            )
                        ),
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new Shading { Fill = "FFFFFF" }),
                                WordServiceSetting.BoldParagraph("-")
                            )
                        )
                    );
                    body.AppendChild(legalBasisTable);
                }



                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("4.คู่สัญญารับทราบและตกลงว่า แต่ละฝ่ายต่างเป็นผู้ควบคุมข้อมูลส่วนบุคคลในส่วนของข้อมูลส่วนบุคคลที่ตนประมวลผล และต่างอยู่ภายใต้บังคับในการปฏิบัติตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลในบทบัญญัติที่เกี่ยวข้องกับผู้ควบคุมข้อมูลส่วนบุคคลต่างหากจากกัน", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("5.คู่สัญญารับรองและยืนยันว่า ก่อนการแบ่งปันข้อมูลส่วนบุคคลแก่อีกฝ่าย ตนได้ดำเนินการแจ้งข้อมูลที่จำเป็นเกี่ยวกับการแบ่งปันข้อมูลและขอความยินยอมจากเจ้าของข้อมูลส่วนบุคคล และ/หรือ มีฐานกฎหมายหรืออำนาจหน้าที่โดยชอบด้วยกฎหมายให้สามารถเปิดเผยข้อมูลส่วนบุคคลให้อีกฝ่าย และให้อีกฝ่ายสามารถทำการประมวลผลข้อมูลส่วนบุคคลที่ได้รับนั้นตามวัตถุประสงค์ที่ได้ตกลงกันอย่างถูกต้องตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลแล้ว", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("6." + "คู่สัญญารับรองว่า คู่สัญญาฝ่ายที่แบ่งปันข้อมูลส่วนบุคคล จะไม่ถูกจำกัดสิทธิ ยับยั้งหรือมีข้อห้ามใด ๆ ในการ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("6.1.ประมวลผลข้อมูลส่วนบุคคลที่ตนเป็นฝ่ายแบ่งปัน ภายใต้วัตถุประสงค์ที่กำหนดในข้อตกลงฉบับนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("6.2.แบ่งปันส่วนบุคคลไปยังคู่สัญญาอีกฝ่ายเพื่อการปฏิบัติหน้าที่ตามข้อตกลงฉบับนี้", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("7.คู่สัญญาจะทำการประมวลผลข้อมูลส่วนบุคคลที่รับมาจากอีกฝ่ายเพียงเท่าที่จำเป็น เพื่อให้บรรลุวัตถุประสงค์ที่ได้กำหนดในข้อ 2 ของข้อตกลงฉบับนี้และแต่ละฝ่ายจะไม่ประมวลผลข้อมูล เพื่อวัตถุประสงค์อื่นเว้นแต่ได้รับความยินยอมจากเจ้าของข้อมูลส่วนบุคคล หรือเป็นความจำเป็นเพื่อปฏิบัติตามกฎหมายเท่านั้น", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("8.คู่สัญญารับรองว่าจะควบคุมดูแลให้เจ้าหน้าที่ และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ปฏิบัติหน้าที่ในการประมวลผลข้อมูลส่วนบุคคลที่ได้รับจากอีกฝ่ายภายใต้ข้อตกลงฉบับนี้ รักษาความลับและปฏิบัติตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลอย่างเคร่งครัด และดำเนินการประมวลผลข้อมูลส่วนบุคคลเพื่อวัตถุประสงค์ตามข้อตกลงฉบับนี้เท่านั้น โดยจะไม่ทำซ้ำ คัดลอก ทำสำเนา บันทึกภาพข้อมูลส่วนบุคคลไม่ว่าทั้งหมดหรือแต่บางส่วนเป็นอันขาด เว้นแต่เป็นไปตามเงื่อนไขของสัญญาหลัก หรือกฎหมายที่เกี่ยวข้องจะระบุหรือบัญญัติไว้เป็นประการอื่น", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("9.คู่สัญญารับรองว่าจะกำหนดให้การเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้ ถูกจำกัดเฉพาะเจ้าหน้าที่ และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ได้รับมอบหมาย มีหน้าที่เกี่ยวข้องหรือมีความจำเป็นในการเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้เท่านั้น", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("10.คู่สัญญาฝ่ายที่รับข้อมูลจะไม่เปิดเผยข้อมูลส่วนบุคคลที่ได้รับจากฝ่ายที่โอนข้อมูล แก่บุคคลของคู่สัญญาฝ่ายที่รับข้อมูลที่ไม่มีอำนาจหน้าที่ที่เกี่ยวข้องในการประมวลผล หรือบุคคลภายนอกใด ๆ เว้นแต่ที่มีความจำเป็นต้องกระทำตามหน้าที่ในสัญญาหลัก ข้อตกลงฉบับนี้หรือเพื่อปฏิบัติตามกฎหมายที่ใช้บังคับ หรือ ที่ได้รับความยินยอมเป็นลายลักษณ์อักษรจากคู่สัญญาฝ่ายที่โอนข้อมูลก่อน", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("11.คู่สัญญาจัดให้มีและคงไว้ซึ่งมาตรการรักษาความปลอดภัยสำหรับการประมวลผล ข้อมูลที่มีความเหมาะสม" +
                    " ทั้งในเชิงองค์กรและเชิงเทคนิคตามที่คณะกรรมการคุ้มครองข้อมูลส่วนบุคคลได้ประกาศกำหนดและหรือตามมาตรฐานสากล โดยคำนึงถึงลักษณะ ขอบเขต และวัตถุประสงค์ของการประมวลผลข้อมูล เพื่อคุ้มครองข้อมูลส่วนบุคคลจากความเสี่ยงอันเกี่ยวเนื่องกับการประมวลผลข้อมูลส่วนบุคคล เช่น ความเสียหายอันเกิดจากการละเมิด อุบัติเหตุ ลบ ทำลาย สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้ เปิดเผย หรือโอนข้อมูลส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("12. กรณีที่คู่สัญญาฝ่ายหนึ่งฝ่ายใด พบพฤติการณ์ที่มีลักษณะที่กระทบต่อการรักษาความปลอดภัยของข้อมูลส่วนบุคคลที่แบ่งปันกันภายใต้ข้อตกลงฉบับนี้ ซึ่งอาจก่อให้เกิดความเสียหายจากการละเมิด " +
                    "อุบัติเหตุ ลบ ทำลาย สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้ เปิดเผยหรือโอนข้อมูลส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย คู่สัญญาฝ่ายที่พบเหตุดังกล่าวจะดำเนินการแจ้งให้คู่สัญญาอีกฝ่ายทราบโดยทันทีภายในเวลาไม่เกิน " + result.RetentionPeriodDays + " ชั่วโมง", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("13. การแจ้งถึงเหตุการละเมิดข้อมูลส่วนบุคคลที่เกิดขึ้นภายใต้ข้อตกลงนี้ คู่สัญญาแต่ละฝ่ายจะใช้มาตรการตามที่เห็นสมควรในการระบุถึงสาเหตุของการละเมิด " +
                    "และป้องกันปัญหาดังกล่าวมิให้เกิดซ้ำ และจะให้ข้อมูลแก่อีกฝ่ายภายใต้ขอบเขตที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลได้กำหนด ดังต่อไปนี้", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("13.1รายละเอียดของลักษณะและผลกระทบที่อาจเกิดขึ้นของการละเมิด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("13.2มาตรการที่ถูกใช้เพื่อลดผลกระทบของการละเมิด", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("13.3ประเภทของข้อมูลส่วนบุคคลและเจ้าของข้อมูลส่วนบุคคลที่ถูกละเมิด หากมีปรากฏ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_3Tabs("13.4ข้อมูลอื่น ๆ เกี่ยวข้องกับการละเมิด", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("14.คู่สัญญาตกลงจะให้ความช่วยเหลืออย่างสมเหตุสมผลแก่อีกฝ่าย เพื่อปฏิบัติตามกฎหมายคุ้มครองข้อมูลที่ใช้บังคับในการตอบสนองต่อข้อเรียกร้องใด ๆ " +
                    "ที่สมเหตุสมผลจากการใช้สิทธิต่าง ๆ ภายใต้กฎหมายคุ้มครองข้อมูลส่วนบุคคลโดยเจ้าของข้อมูลส่วนบุคคล โดยพิจารณาถึงลักษณะการประมวลผล ภาระหน้าที่ภายใต้กฎหมายคุ้มครองข้อมูลส่วนบุคคลที่ใช้บังคับ และข้อมูลส่วนบุคคลที่แต่ละฝ่ายประมวลผล", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ทั้งนี้ กรณีที่เจ้าของข้อมูลส่วนบุคคลยื่นคำร้องขอใช้สิทธิดังกล่าวต่อคู่สัญญาฝ่ายหนึ่งฝ่ายใด เพื่อใช้สิทธิในข้อมูลส่วนบุคคลที่อยู่ในความรับผิดชอบหรือได้รับมาจากอีกฝ่าย " +
                    "คู่สัญญาฝ่ายที่ได้รับคำร้องจะต้องดำเนินการแจ้งและส่งคำร้องดังกล่าวให้แก่คู่สัญญาซึ่งเป็นฝ่ายโอนข้อมูลโดยทันที โดยคู่สัญญาฝ่ายที่รับคำร้องนั้นจะต้องแจ้งให้เจ้าของข้อมูลส่วนบุคคลทราบถึงการจัดการตามคำขอหรือข้อร้องเรียนของเจ้าของข้อมูลส่วนบุคคลนั้นด้วย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("15.หากคู่สัญญาฝ่ายหนึ่งฝ่ายใดมีความจำเป็นจะต้องเปิดเผยข้อมูลส่วนบุคคลที่ได้รับจากอีกฝ่ายไปยังต่างประเทศ การส่งออกซึ่งข้อมูลส่วนบุคคลดังกล่าวจะต้อง" +
                    "ได้รับปกป้องตามมาตรฐานการส่งข้อมูลระหว่างประเทศตามที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลของประเทศที่ส่งข้อมูลไปนั้นกำหนด ทั้งนี้ คู่สัญญาทั้งสองฝ่ายตกลงที่จะเข้าทำสัญญาใด ๆ ที่จำเป็นต่อการปฏิบัติตามกฎหมายที่ใช้บังคับกับการโอนข้อมูล", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("16.คู่สัญญาแต่ละฝ่ายอาจใช้ผู้ประมวลผลข้อมูลส่วนบุคคล เพื่อทำการประมวลผลข้อมูลส่วนบุคคลที่โอนและรับโอน " +
                    "โดยคู่สัญญาฝ่ายนั้นจะต้องทำสัญญากับผู้ประมวลผลข้อมูลเป็นลายลักษณ์อักษร " +
                    "ซึ่งสัญญาดังกล่าวจะต้องมีเงื่อนไขในการคุ้มครองข้อมูลส่วนบุคคลที่โอนและรับโอนไม่น้อยไปกว่าเงื่อนไขที่ได้ระบุไว้ในข้อตกลงฉบับนี้ " +
                    "และเงื่อนไขทั้งหมดต้องเป็นไปตามที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลกำหนด เพื่อหลีกเลี่ยงข้อสงสัย หากคู่สัญญาฝ่ายหนึ่งฝ่ายใดได้ว่าจ้างหรือมอบหมายผู้ประมวลผลข้อมูลส่วนบุคคล " +
                    "คู่สัญญาฝ่ายนั้นยังคงต้องมีความรับผิดต่ออีกฝ่ายสำหรับการกระทำการหรือละเว้นกระทำการใด ๆ ของผู้ประมวลผลข้อมูลส่วนบุคคลนั้น", null, "32"));


                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("17.เว้นแต่กฎหมายที่เกี่ยวข้องจะบัญญัติไว้เป็นประการอื่น คู่สัญญาจะทำการลบหรือทำลายข้อมูลส่วนบุคคลที่ตนได้รับจากอีกฝ่ายภายใต้ข้อตกลงฉบับนี้ภายใน " + result.IncidentNotifyPeriod + " วัน " +
                    "นับแต่วันที่ดำเนินการประมวลผลตามวัตถุประสงค์ภายใต้ข้อตกลงฉบับนี้เสร็จสิ้น หรือวันที่คู่สัญญาได้ตกลงเป็นลายลักษณ์อักษรให้ยกเลิกสัญญาหลักแล้วแต่กรณีใดจะเกิดขึ้นก่อน", null, "32"));

                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("18.คู่สัญญาแต่ละฝ่ายจะต้องชดใช้ความเสียหายให้แก่อีกฝ่ายในค่าปรับ ความสูญหายหรือเสียหายใด ๆ ที่เกิดขึ้นกับฝ่ายที่ไม่ได้ผิดเงื่อนไข อันเนื่องมาจากการฝ่าฝืนข้อตกลงฉบับนี้" +
                    " แม้ว่าจะมีข้อจำกัดความรับผิดภายใต้สัญญาหลักก็ตาม", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("19.หน้าที่และความรับผิดของคู่สัญญาในการปฏิบัติตามข้อตกลงฉบับนี้จะสิ้นสุดลงนับแต่วันที่การดำเนินการตามสัญญาหลักเสร็จสิ้นลง หรือ วันที่คู่สัญญาได้ตกลงเป็นลายลักษณ์อักษรให้ยกเลิกสัญญาหลัก" +
                    " แล้วแต่กรณีใดจะเกิดขึ้นก่อน อย่างไรก็ดี การสิ้นผลลงของข้อตกลงฉบับนี้ ไม่กระทบต่อหน้าที่ของคู่สัญญาแต่ละฝ่ายในการลบหรือทำลายข้อมูลส่วนบุคคลตามที่ได้กำหนดในข้อ 17 ของข้อตกลงฉบับนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ในกรณีที่ข้อตกลง คำรับรอง การเจรจา หรือข้อผูกพันใดที่คู่สัญญามีต่อกันไม่ว่าด้วยวาจาหรือเป็นลายลักษณ์อักษรใดขัดหรือแย้งกับข้อตกลงที่ระบุในข้อตกลงฉบับนี้ ให้ใช้ข้อความตามข้อตกลงฉบับนี้บังคับ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ทั้งสองฝ่ายได้อ่านและเข้าใจข้อความโดยละเอียดตลอดแล้ว เพื่อเป็นหลักฐานแห่งการนี้ ทั้งสองฝ่ายจึงได้ลงนามไว้เป็นหลักฐานต่อหน้าพยาน ณ วัน เดือน ปี ที่ระบุข้างต้น", null, "32"));
                //body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("xxx", null, "32"));
                //body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("xxx", null, "32"));
                //body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("xxx", null, "32"));
                //body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("xxx", null, "32"));
                //body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("xxx", null, "32"));
                //body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("xxx", null, "32"));
                //body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("xxx", null, "32"));

                // --- 6. Signature lines ---
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                var signatureTable = new Table(
          new TableProperties(
              new TableWidth { Width = "10000", Type = TableWidthUnitValues.Dxa }, // ขยายตาราง
              new TableJustification { Val = TableRowAlignmentValues.Center },      // จัดตรงกลาง
              new TableBorders(
                  new TopBorder { Val = BorderValues.None },
                  new BottomBorder { Val = BorderValues.None },
                  new LeftBorder { Val = BorderValues.None },
                  new RightBorder { Val = BorderValues.None },
                  new InsideHorizontalBorder { Val = BorderValues.None },
                  new InsideVerticalBorder { Val = BorderValues.None }
              )
          ),
          // แถวลายเซ็นหลัก
          new TableRow(
              new TableCell(
                  new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                  WordServiceSetting.CenteredParagraph("ลงชื่อ " + (string.IsNullOrWhiteSpace(result.OSMEP_Signer) ? "รอผู้ลงนาม" : result.OSMEP_Signer)),

                  WordServiceSetting.CenteredParagraph("(............................................................)"),
                  WordServiceSetting.CenteredParagraph("ผู้อำนวยการสำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม หรือผู้ที่ผู้อำนวยการมอบหมาย")
              ),
              new TableCell(
                  new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                                WordServiceSetting.CenteredParagraph("ลงชื่อ " + (string.IsNullOrWhiteSpace(result.Contract_Signer) ? "รอผู้ลงนาม" : result.OSMEP_Signer)),

                  WordServiceSetting.CenteredParagraph("(............................................................)"),
                  WordServiceSetting.CenteredParagraph("............................................................")
              )
          ),
          // แถวพยาน
          new TableRow(
              new TableCell(
                  new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                  WordServiceSetting.CenteredParagraph("ลงชื่อ " + (string.IsNullOrWhiteSpace(result.OSMEP_Witness) ? "รอผู้ลงนาม" : result.OSMEP_Witness)),
                  WordServiceSetting.CenteredParagraph("พยาน"),
                  WordServiceSetting.CenteredParagraph("(............................................................)"),
                  WordServiceSetting.CenteredParagraph("............................................................")
              ),
              new TableCell(
                  new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                  WordServiceSetting.CenteredParagraph("ลงชื่อ " + (string.IsNullOrWhiteSpace(result.Contract_Witness) ? "รอผู้ลงนาม" : result.Contract_Witness)),
                  WordServiceSetting.CenteredParagraph("พยาน"),
                  WordServiceSetting.CenteredParagraph("(............................................................)"),
                  WordServiceSetting.CenteredParagraph("............................................................")
              )
          )
      );
                body.AppendChild(signatureTable);
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                body.AppendChild(WordServiceSetting.EmptyParagraph());

                WordServiceSetting.AddHeaderWithPageNumber(mainPart, body);
            }
            stream.Position = 0;
            return stream.ToArray();
        }
        catch (Exception ex)

        {
            throw new Exception("Error in WordEContract_DataPersonalService.OnGetWordContact_DataPersonalService: " + ex.Message);
        }

    }

    #endregion  4.1.1.2.6.บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล

    #region PDSA
    public async Task<byte[]> OnGetWordContact_DataPersonalService_ToPDF(string id,string typeContact)
    {
        try
        {
            var result = await _e.GetPDSAAsync(id);
            var rLe = await _e.GetPDSA_LegalBasisSharingAsync(id);
            var rSd = await _e.GetPDSA_Shared_DataAsync(id);
            var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");
            var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");

string datestring = CommonDAO.ToThaiDateStringCovert(result.Master_Contract_Sign_Date ?? DateTime.Now);            string logoBase64 = "";
            if (System.IO.File.Exists(logoPath))
            {
                var bytes = System.IO.File.ReadAllBytes(logoPath);
                logoBase64 = Convert.ToBase64String(bytes);
            }

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
          
        }}
         body {{
            font-size: 22px;
            font-family: 'THSarabunNew', Arial, sans-serif;
         
        }}
        .t-16 {{
            font-size: 1.8em;
        }}
        .t-18 {{
            font-size: 2.0em;
        }}
        .t-22 {{
            font-size: 2.2em;
        }}
        .tab1 {{ text-indent: 48px;  word-break: break-all;  }}
        .tab2 {{ text-indent: 96px;  word-break: break-all; }}
        .tab3 {{ text-indent: 144px;  word-break: break-all; }}
        .tab4 {{ text-indent: 192px;  word-break: break-all;}}
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
        .table {{ width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 28pt; }}
        .table th, .table td {{ border: 1px solid #000; padding: 8px; }}

        .sign-double {{ display: flex; }}
        .text-center-right-brake {{
            margin-left: 50%;
            word-break: break-all;
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
    <!-- Logo Row -->
        <table class='contract-table'>
        <tr>
            <td style='width:60%; text-align:left;'>
                <img src='data:image/jpeg;base64,{logoBase64}' class='logo-img' />
            </td>
            <td style='width:40%; text-align:center;'>
                <img src='data:image/jpeg;base64,{logoBase64}' class='logo-img' />
            </td>
        </tr>
    </table>
</br>

    <!-- Titles -->
    <div class='t-22 text-center'><b>ข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล</b></div>
    <div class='t-18 text-center'><b>(Personal Data Sharing Agreement)</b></div>
    <div class='t-18 text-center'><b>ระหว่าง</b></div>
    <div class='t-18 text-center'><b>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม กับ {result.ContractPartyName}</b></div>
    <div class='t-18 text-center'>---------------------------------------------</div>
    <!-- Paragraphs -->
    <p class='tab3 t-16'>ข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล (“ข้อตกลง”) ฉบับนี้ทำขึ้น เมื่อ {datestring} ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</p>
    <p class='tab3 t-16'>โดยที่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ... รวมเรียกว่า “คู่สัญญา”</p>
    <p class='tab3 t-16'>เพื่อให้บรรลุวัตถุประสงค์ภายใต้ความตกลงของสัญญาหลัก ...</p>
    <p class='tab3 t-16'>1. คู่สัญญารับทราบว่า ข้อมูลส่วนบุคคล หมายถึง ข้อมูลเกี่ยวกับบุคคลธรรมดา ซึ่งทำให้สามารถระบุตัวบุคคลนั้นได้ไม่ว่าทางตรงหรือทางอ้อม โดยคู่สัญญาแต่ละฝ่าย จะดำเนินการตามที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลกำหนด 
</br>เพื่อคุ้มครองให้การประมวลผลข้อมูลส่วนบุคคลเป็นไปอย่างเหมาะสมและถูกต้องตามกฎหมาย</p>
    <p class='tab3 t-16'>2. ข้อมูลส่วนบุคคลที่คู่สัญญาแบ่งปันกัน คู่สัญญาแต่ละฝ่ายตกลงแบ่งปันข้อมูลส่วนบุคคลดังรายการต่อไปนี้แก่คู่สัญญาอีกฝ่าย</p>
    <!-- Table: ข้อมูลส่วนบุคคลที่แบ่งปันโดย สสว. -->
    <table class='table t-16'>
        <tr>
            <th>ข้อมูลส่วนบุคคลที่แบ่งปันโดย สสว.</th>
            <th>วัตถุประสงค์ในการแบ่งปันข้อมูลส่วนบุคคล</th>
        </tr>
{string.Join("", rSd.Where(e => e.Owner == "OSMEP").Select(item => $"<tr><td>{item.Detail ?? "-"}</td><td>{item.Objective ?? "-"}</td></tr>"))}
    </table>
    <!-- Table: ข้อมูลส่วนบุคคลที่แบ่งปันโดยคู่สัญญา -->
   <table class='table t-16'>
        <tr>
            <th>ข้อมูลส่วนบุคคลที่แบ่งปันโดย {result.ContractPartyName}</th>
            <th>วัตถุประสงค์ในการแบ่งปันข้อมูลส่วนบุคคล</th>
        </tr>
      {string.Join("", rSd.Where(e => e.Owner == "CP").Select(item => $"<tr><td>{item.Detail ?? "-"}</td><td>{item.Objective ?? "-"}</td></tr>"))}
    </table>
    <!-- Table: ฐานกฎหมายในการแบ่งปันข้อมูลส่วนบุคคล -->
    <p class='tab3 t-16'>3. ฐานกฎหมายในการแบ่งปันข้อมูลส่วนบุคคล ภายใต้วัตถุประสงค์ที่ระบุในข้อ 2 คู่สัญญาแต่ละฝ่ายมีฐานกฎหมายตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลดังต่อไปนี้ ในการแบ่งปันข้อมูลส่วนบุคคลแก่คู่สัญญาอีกฝ่าย (แต่ละฝ่ายอาจ</br>ใช้ฐานกฎหมายที่ต่างกันในการแบ่งปันข้อมูลส่วนบุคคล)</p>
   <table class='table t-16'>
        <tr>
            <th>ฐานกฎหมายของ สสว.</th>
        </tr>
   {string.Join("", rLe.Where(e => e.Owner == "OSMEP").Select(item => $"<tr><td>{item.Detail ?? "-"}</td></tr>"))}
    </table>
   <table class='table t-16'>
        <tr>
            <th>ฐานกฎหมายของ {result.ContractPartyName}</th>
        </tr>
    {string.Join("", rLe.Where(e => e.Owner == "CP").Select(item => $"<tr><td>{item.Detail ?? "-"}</td></tr>"))}
    </table>
<!-- No file path since this is a template snippet -->
 <p class='tab3 t-16'>4. คู่สัญญารับทราบและตกลงว่า แต่ละฝ่ายต่างเป็นผู้ควบคุมข้อมูลส่วนบุคคลในส่วนของข้อมูลส่วนบุคคลที่ตนประมวลผล และต่างอยู่ภายใต้บังคับในการปฏิบัติตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลในบทบัญญัติที่เกี่ยวข้องกับผู้ควบคุมข้อมูลส่วนบุคคลต่างหากจากกัน</p>
 <p class='tab3 t-16'>5. คู่สัญญารับรองและยืนยันว่า ก่อนการแบ่งปันข้อมูลส่วนบุคคลแก่อีกฝ่าย ตนได้ดำเนินการแจ้งข้อมูลที่</br>จำเป็นเกี่ยวกับการแบ่งปันข้อมูลและขอความยินยอมจากเจ้าของข้อมูลส่วนบุคคล และ/หรือ 
มีฐานกฎหมายหรืออำนาจหน้าที่โดยชอบด้วยกฎหมายให้สามารถเปิดเผยข้อมูลส่วนบุคคลให้อีกฝ่าย และให้อีกฝ่ายสามารถทำการประมวลผล</br>ข้อมูลส่วนบุคคลที่ได้รับนั้นตามวัตถุประสงค์ที่ได้ตกลงกันอย่างถูกต้องตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลแล้ว</p>
 <p class='tab3 t-16'>6. คู่สัญญารับรองว่า คู่สัญญาฝ่ายที่แบ่งปันข้อมูลส่วนบุคคล จะไม่ถูกจำกัดสิทธิ ยับยั้งหรือมีข้อห้ามใด ๆ</br> ในการ</p>
 <p class='tab3 t-16'>6.1.ประมวลผลข้อมูลส่วนบุคคลที่ตนเป็นฝ่ายแบ่งปัน ภายใต้วัตถุประสงค์ที่กำหนดในข้อตกลงฉบับนี้</p>
 <p class='tab3 t-16'>6.2.แบ่งปันส่วนบุคคลไปยังคู่สัญญาอีกฝ่ายเพื่อการปฏิบัติหน้าที่ตามข้อตกลงฉบับนี้</p>
 <p class='tab3 t-16'>7. คู่สัญญาจะทำการประมวลผลข้อมูลส่วนบุคคลที่รับมาจากอีกฝ่ายเพียงเท่าที่จำเป็น เพื่อให้บรรลุวัตถุ</br>ประสงค์ที่ได้กำหนดในข้อ 2 ของข้อตกลงฉบับนี้และแต่ละฝ่ายจะไม่ประมวลผลข้อมูล เพื่อวัตถุประสงค์อื่นเว้นแต่ได้รับความยินยอมจากเจ้าของข้อมูลส่วนบุคคล หรือเป็นความจำเป็นเพื่อปฏิบัติตามกฎหมายเท่านั้น</p>
 <p class='tab3 t-16'>8. คู่สัญญารับรองว่าจะควบคุมดูแลให้เจ้าหน้าที่ และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ปฏิบัติหน้าที่ใน</br>การประมวลผลข้อมูลส่วนบุคคลที่ได้รับจากอีกฝ่ายภายใต้ข้อตกลงฉบับนี้
รักษาความลับและปฏิบัติตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลอย่างเคร่งครัด และดำเนินการประมวลผลข้อมูลส่วนบุคคลเพื่อวัตถุประสงค์ตามข้อตกลงฉบับนี้เท่านั้น </br>โดยจะไม่ทำซ้ำ คัดลอก ทำสำเนา บันทึกภาพข้อมูลส่วนบุคคลไม่ว่าทั้งหมดหรือแต่บางส่วนเป็นอันขาด
เว้นแต่เป็นไป</br>ตามเงื่อนไขของสัญญาหลัก หรือกฎหมายที่เกี่ยวข้องจะระบุหรือบัญญัติไว้เป็นประการอื่น</p>
 <p class='tab3 t-16'>9. คู่สัญญารับรองว่าจะกำหนดให้การเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้ ถูกจำกัดเฉพาะเจ้าหน้าที่</br> และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ได้รับมอบหมาย มีหน้าที่เกี่ยวข้องหรือมีความจำเป็นในการเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้เท่านั้น</p>
 <p class='tab3 t-16'>10. คู่สัญญาฝ่ายที่รับข้อมูลจะไม่เปิดเผยข้อมูลส่วนบุคคลที่ได้รับจากฝ่ายที่โอนข้อมูล 
แก่บุคคลของคู่สัญญาฝ่ายที่รับข้อมูลที่ไม่มีอำนาจหน้าที่ที่เกี่ยวข้องในการประมวลผล หรือบุคคลภายนอกใด ๆ
เว้นแต่ที่มีความจำเป็นต้องกระทำตามหน้าที่ในสัญญาหลัก ข้อตกลงฉบับนี้หรือเพื่อปฏิบัติตามกฎหมายที่ใช้บังคับ หรือ ที่ได้รับความยินยอมเป็นลายลักษณ์</br>อักษรจากคู่สัญญาฝ่ายที่โอนข้อมูลก่อน</p>
 <p class='tab3 t-16'>11. คู่สัญญาจัดให้มีและคงไว้ซึ่งมาตรการรักษาความปลอดภัยสำหรับการประมวลผล ข้อมูลที่มีความเหมาะสม 
ทั้งในเชิงองค์กรและเชิงเทคนิคตามที่คณะกรรมการคุ้มครองข้อมูลส่วนบุคคลได้ประกาศกำหนดและหรือตามมาตรฐานสากล 
โดยคำนึงถึงลักษณะ ขอบเขต และวัตถุประสงค์ของการประมวลผลข้อมูล เพื่อคุ้มครองข้อมูลส่วนบุคคลจากความเสี่ยง</br>อันเกี่ยวเนื่องกับการประมวลผลข้อมูลส่วนบุคคล เช่น ความเสียหายอันเกิดจากการละเมิด อุบัติเหตุ ลบ ทำลาย สูญหาย </br>เปลี่ยนแปลง แก้ไข เข้าถึง ใช้ เปิดเผย หรือโอนข้อมูลส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย</p>
 <p class='tab3 t-16'>12. กรณีที่คู่สัญญาฝ่ายหนึ่งฝ่ายใด พบพฤติการณ์ที่มีลักษณะที่กระทบต่อการรักษาความปลอดภัยของ</br>ข้อมูลส่วนบุคคลที่แบ่งปันกันภายใต้ข้อตกลงฉบับนี้ ซึ่งอาจก่อให้เกิดความเสียหายจากการละเมิด อุบัติเหตุ ลบ ทำลาย </br>สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้
เปิดเผยหรือโอนข้อมูลส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย คู่สัญญาฝ่ายที่พบเหตุดังกล่าว</br>จะดำเนินการแจ้งให้คู่สัญญาอีกฝ่ายทราบโดยทันทีภายในเวลาไม่เกิน {result.RetentionPeriodDays} ชั่วโมง</p>

 <p class='tab3 t-16'>13 การแจ้งถึงเหตุการละเมิดข้อมูลส่วนบุคคลที่เกิดขึ้นภายใต้ข้อตกลงนี้ คู่สัญญาแต่ละฝ่ายจะใช้มาตรการตามที่เห็นสมควรในการระบุถึงสาเหตุของการละเมิด
และป้องกันปัญหาดังกล่าวมิให้เกิดซ้ำ และจะให้ข้อมูลแก่อีกฝ่ายภายใต้ขอบเขตที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลได้กำหนด ดังต่อไปนี้</p>
 <p class='tab3 t-16'>13.1 รายละเอียดของลักษณะและผลกระทบที่อาจเกิดขึ้นของการละเมิด</p>
 <p class='tab3 t-16'>13.2 มาตรการที่ถูกใช้เพื่อลดผลกระทบของการละเมิด</p>
 <p class='tab3 t-16'>13.3 ประเภทของข้อมูลส่วนบุคคลและเจ้าของข้อมูลส่วนบุคคลที่ถูกละเมิด หากมีปรากฏ</p>
 <p class='tab3 t-16'>13.4 ข้อมูลอื่น ๆ เกี่ยวข้องกับการละเมิด</p>
 <p class='tab3 t-16'>14. 
คู่สัญญาตกลงจะให้ความช่วยเหลืออย่างสมเหตุสมผลแก่อีกฝ่าย เพื่อปฏิบัติตามกฎหมายคุ้มครอง
</br>ข้อมูลที่ใช้บังคับในการตอบสนองต่อข้อเรียกร้องใด ๆ ที่สมเหตุสมผลจากการใช้สิทธิต่าง ๆ ภายใต้กฎหมายคุ้มครอง
</br>ข้อมูลส่วนบุคคลโดยเจ้าของข้อมูลส่วนบุคคล โดยพิจารณาถึงลักษณะการประมวลผล ภาระหน้าที่ภายใต้กฎหมาย
</br>คุ้มครองข้อมูลส่วนบุคคลที่ใช้บังคับ และข้อมูลส่วนบุคคลที่แต่ละฝ่ายประมวลผล</p>
 <p class='tab3 t-16'>
ทั้งนี้ กรณีที่เจ้าของข้อมูลส่วนบุคคลยื่นคำร้องขอใช้สิทธิดังกล่าวต่อคู่สัญญาฝ่ายหนึ่งฝ่ายใด เพื่อใช้สิทธิใน
</br>ข้อมูลส่วนบุคคลที่อยู่ในความรับผิดชอบหรือได้รับมาจากอีกฝ่าย คู่สัญญาฝ่ายที่ได้รับคำร้องจะต้องดำเนินการแจ้งและ
</br>ส่งคำร้องดังกล่าวให้แก่คู่สัญญาซึ่งเป็นฝ่ายโอนข้อมูลโดยทันที โดยคู่สัญญาฝ่ายที่รับคำร้องนั้นจะต้องแจ้งให้เจ้าของ
</br>ข้อมูลส่วนบุคคลทราบถึงการจัดการตามคำขอหรือข้อร้องเรียนของเจ้าของข้อมูลส่วนบุคคลนั้นด้วย</p>
 <p class='tab3 t-16'>
15. หากคู่สัญญาฝ่ายหนึ่งฝ่ายใดมีความจำเป็นจะต้องเปิดเผยข้อมูลส่วนบุคคลที่ได้รับจากอีกฝ่ายไปยังต่าง
</br>ประเทศ การส่งออกซึ่งข้อมูลส่วนบุคคลดังกล่าวจะต้องได้รับปกป้องตามมาตรฐานการส่งข้อมูลระหว่างประเทศตามที่
</br>กฎหมายคุ้มครองข้อมูลส่วนบุคคลของประเทศที่ส่งข้อมูลไปนั้นกำหนด ทั้งนี้ คู่สัญญาทั้งสองฝ่ายตกลงที่จะเข้าทำ
</br>สัญญาใด ๆ ที่จำเป็นต่อการปฏิบัติตามกฎหมายที่ใช้บังคับกับการโอนข้อมูล</p>
 <p class='tab3 t-16'>
16. คู่สัญญาแต่ละฝ่ายอาจใช้ผู้ประมวลผลข้อมูลส่วนบุคคล เพื่อทำการประมวลผลข้อมูลส่วนบุคคลที่โอนและ
</br>รับโอน โดยคู่สัญญาฝ่ายนั้นจะต้องทำสัญญากับผู้ประมวลผลข้อมูลเป็นลายลักษณ์อักษร ซึ่งสัญญาดังกล่าวจะต้องมีเงื่อนไข
</br>ในการคุ้มครองข้อมูลส่วนบุคคลที่โอนและรับโอนไม่น้อยไปกว่าเงื่อนไขที่ได้ระบุไว้ในข้อตกลงฉบับนี้ และเงื่อนไข
</br>ทั้งหมดต้องเป็นไปตามที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลกำหนด เพื่อหลีกเลี่ยงข้อสงสัย หากคู่สัญญาฝ่ายหนึ่งฝ่ายใด
</br>ได้ว่าจ้างหรือมอบหมายผู้ประมวลผลข้อมูลส่วนบุคคล คู่สัญญาฝ่ายนั้นยังคงต้องมีความรับผิดต่ออีกฝ่ายสำหรับการ
</br>กระทำการหรือละเว้นกระทำการใด ๆ ของผู้ประมวลผลข้อมูลส่วนบุคคลนั้น</p>
 <p class='tab3 t-16'>
17. เว้นแต่กฎหมายที่เกี่ยวข้องจะบัญญัติไว้เป็นประการอื่นคู่สัญญาจะทำการลบหรือทำลายข้อมูลส่วนบุคคล
</br>ที่ตนได้รับจากอีกฝ่ายภายใต้ข้อตกลงฉบับนี้ภายใน {result.IncidentNotifyPeriod} วัน นับแต่วันที่ดำเนินการประมวลผลตามวัตถุประสงค์
</br>ภายใต้ข้อตกลงฉบับนี้เสร็จสิ้น หรือวันที่คู่สัญญาได้ตกลงเป็นลายลักษณ์อักษรให้ยกเลิกสัญญาหลัก
</br>แล้วแต่กรณีใดจะเกิดขึ้นก่อน</p>
 <p class='tab3 t-16'>18. คู่สัญญาแต่ละฝ่ายจะต้องชดใช้ความเสียหายให้แก่อีกฝ่ายในค่าปรับ ความสูญหายหรือเสียหายใด ๆ ที่เกิดขึ้นกับฝ่ายที่ไม่ได้ผิดเงื่อนไข อันเนื่องมาจากการฝ่าฝืนข้อตกลงฉบับนี้ แม้ว่าจะมีข้อจำกัดความรับผิดภายใต้สัญญาหลักก็ตาม</p>
 <p class='tab3 t-16'>
19. หน้าที่และความรับผิดของคู่สัญญาในการปฏิบัติตามข้อตกลงฉบับนี้จะสิ้นสุดลงนับแต่วันที่การดำเนิน
</br>การตามสัญญาหลักเสร็จสิ้นลง หรือ วันที่คู่สัญญาได้ตกลงเป็นลายลักษณ์อักษรให้ยกเลิกสัญญาหลัก แล้วแต่กรณีใดจะเกิดขึ้นก่อน อย่างไรก็ดี การสิ้นผลลงของข้อตกลงฉบับนี้ ไม่กระทบต่อหน้าที่ของคู่สัญญาแต่ละฝ่ายในการลบหรือทำลายข้อมูลส่วนบุคคลตามที่ได้กำหนดในข้อ 17 ของข้อตกลงฉบับนี้</p>
 <p class='tab3 t-16'>ในกรณีที่ข้อตกลง คำรับรอง การเจรจา หรือข้อผูกพันใดที่คู่สัญญามีต่อกันไม่ว่าด้วยวาจาหรือเป็น
</br>ลายลักษณ์อักษรใดขัดหรือแย้งกับข้อตกลงที่ระบุในข้อตกลงฉบับนี้ ให้ใช้ข้อความตามข้อตกลงฉบับนี้บังคับ</p>
 <p class='tab3 t-16'>ทั้งสองฝ่ายได้อ่านและเข้าใจข้อความโดยละเอียดตลอดแล้ว เพื่อเป็นหลักฐานแห่งการนี้ ทั้งสองฝ่ายจึงได้</br>ลงนามไว้เป็นหลักฐานต่อหน้าพยาน ณ วัน เดือน ปี ที่ระบุข้างต้น</p>
    <!-- Signature Table -->

</br>
</br>
{signatoryHtml}
</body>
</html>
";
            var doc = new DinkToPdf.HtmlToPdfDocument()
            {
                GlobalSettings = {
                PaperSize = DinkToPdf.PaperKind.A4,
                Orientation = DinkToPdf.Orientation.Portrait,
                Margins = new DinkToPdf.MarginSettings
                {
                    Top = 20,
                    Bottom = 20,
                    Left = 20,
                    Right = 20
                }
            },
                Objects = {
                new DinkToPdf.ObjectSettings() {
                    HtmlContent = html,
                    FooterSettings = new DinkToPdf.FooterSettings
                    {
                        FontName = "THSarabunNew",
                        FontSize = 6,
                        Line = false,
                        Center = "[page] / [toPage]"
                    }
                }
            }
            };

            var pdfBytes = _pdfConverter.Convert(doc);
            return pdfBytes;
        }
        catch (Exception ex)
        {
            throw new Exception("Error in WordEContract_DataPersonalService.OnGetWordContact_DataPersonalService_ToPDF: " + ex.Message);
        }
    }
    #endregion
}