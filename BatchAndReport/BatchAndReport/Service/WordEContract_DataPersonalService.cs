using BatchAndReport.DAO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading.Tasks;

public class WordEContract_DataPersonalService
{
    private readonly WordServiceSetting _w;
    private readonly Econtract_Report_PDSADAO _e;
    public WordEContract_DataPersonalService(WordServiceSetting ws, Econtract_Report_PDSADAO e)
    {
        _w = ws;
        _e = e;
    }

    #region  4.1.1.2.6.บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล
    public async Task<byte[]> OnGetWordContact_DataPersonalService(string id)
    {
        try { 
        
            var result = await _e.GetPDSAAsync(id);
            var rLe =await _e.GetPDSA_LegalBasisSharingAsync(id);
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
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม กับ "+result.ContractPartyName+" ", "32"));
                body.AppendChild(WordServiceSetting.CenteredBoldParagraph("---------------------------------------------", "32"));

                string datestring = CommonDAO.ToThaiDateStringCovert(result.Master_Contract_Sign_Date ?? DateTime.Now);
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล (“ข้อตกลง”) ฉบับนี้ทำขึ้น เมื่อ "+ datestring + " ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("โดยที่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สสว.” ฝ่ายหนึ่ง ได้ตกลงใน.... (ระบุชื่อบันทึกข้อตกลงความร่วมมือ/สัญญาหลัก) .... สัญญาเลขที่ .......... (ระบุเลขที่บันทึกข้อตกลงความร่วมมือ/สัญญาหลัก).................  ฉบับลงวันที่ ..... (ระบุวันที่ลงนามข้อตกลงความร่วมมือหรือวันทำสัญญาหลัก) .......... ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สัญญาหลัก” กับ ........ (ระบุชื่อคู่สัญญาอีกฝ่าย) ........ ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “..... (ระบุชื่อเรียกคู่สัญญาอีกฝ่าย ......” อีกฝ่ายหนึ่ง รวมเรียกว่า “คู่สัญญา”", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("เพื่อให้บรรลุวัตถุประสงค์ภายใต้ความตกลงของสัญญาหลัก คู่สัญญามีความจำเป็นต้องแบ่งปัน โอน แลกเปลี่ยน หรือเปิดเผย (รวมเรียกว่า “แบ่งปัน”) ข้อมูลส่วนบุคคลที่ตนเก็บรักษาแก่อีกฝ่าย ซึ่งข้อมูลส่วนบุคคลที่แต่ละฝ่าย เก็บรวมรวม ใช้หรือเปิดเผย (รวมเรียกว่า “ประมวลผล”) นั้น แต่ละฝ่ายต่างเป็นผู้ควบคุมข้อมูลส่วนบุคคล ตามกฎหมายที่เกี่ยวข้องกับการคุ้มครองข้อมูลส่วนบุคคล กล่าวคือแต่ละฝ่ายต่างเป็นผู้มีอำนาจตัดสินใจ กำหนดรูปแบบ และกำหนดวัตถุประสงค์ ในการประมวลผลข้อมูลส่วนบุคคลในข้อมูลที่ตนต้องแบ่งปัน ภายใต้ข้อตกลงนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("ด้วยเหตุนี้ คู่สัญญาจึงตกลงจัดทำข้อตกลงฉบับนี้ และให้ถือเป็นส่วนหนึ่งของสัญญาหลัก เพื่อเป็นหลักฐานการแบ่งปันข้อมูลส่วนบุคคลระหว่างคู่สัญญาและเพื่อดำเนินการให้เป็นไปตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒ และกฎหมายอื่น ๆ ที่ออกตามความในพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒ ซึ่งต่อไปในข้อตกลงฉบับนี้ รวมเรียกว่า “กฎหมายคุ้มครองข้อมูลส่วนบุคคล”  ทั้งที่มีผลใช้บังคับอยู่ ณ วันทำข้อตกลงฉบับนี้ และที่จะมีการเพิ่มเติมหรือแก้ไขเปลี่ยนแปลงในภายหลัง โดยมีรายละเอียดดังนี้", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("1. คู่สัญญารับทราบว่า ข้อมูลส่วนบุคคล หมายถึง ข้อมูลเกี่ยวกับบุคคลธรรมดา ซึ่งทำให้สามารถระบุตัวบุคคลนั้นได้ไม่ว่าทางตรงหรือทางอ้อม โดยคู่สัญญาแต่ละฝ่าย จะดำเนินการตามที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลกำหนด เพื่อคุ้มครองให้การประมวลผลข้อมูลส่วนบุคคลเป็นไปอย่างเหมาะสมและถูกต้องตามกฎหมาย", null, "32"));
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("2. ข้อมูลส่วนบุคคลที่คู่สัญญาแบ่งปันกัน คู่สัญญาแต่ละฝ่ายตกลงแบ่งปันข้อมูลส่วนบุคคลดังรายการต่อไปนี้แก่คู่สัญญาอีกฝ่าย", null, "32"));

                // Table: ข้อมูลส่วนบุคคลที่แบ่งปันโดย สสว.
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
                            WordServiceSetting.NormalParagraph("1. (ระบุรายการข้อมูลส่วนบุคคลที่ สสว. แบ่งปันให้คู่สัญญาอีกฝ่าย เช่น ชื่อ นามสกุลของเจ้าหน้าที่ หมายเลขโทรศัพท์ ข้อมูลผู้ใช้งานแอปพลิเคชันทางรัฐ)", null, "32")
                        ),
                        new TableCell(
                            WordServiceSetting.NormalParagraph("1. เพื่อความจำเป็นในการ... (ระบุเหตุผลความจำเป็นในการแบ่งปันข้อมูลส่วนบุคคล ระหว่างคู่สัญญา เช่น เพื่อการเชื่อมโยงแสดงผลข้อมูลในแอปพลิเคชัน)", null, "32")
                        )
                    ),
                    new TableRow(
                        new TableCell(WordServiceSetting.NormalParagraph("2. ...", null, "32")),
                        new TableCell(WordServiceSetting.NormalParagraph("2. ...", null, "32"))
                    ),
                    new TableRow(
                        new TableCell(WordServiceSetting.NormalParagraph("3. ...", null, "32")),
                        new TableCell(WordServiceSetting.NormalParagraph("3. ...", null, "32"))
                    )
                );
                body.AppendChild(infoTable);

                // Table: ฐานกฎหมายในการแบ่งปันข้อมูลส่วนบุคคล
                body.AppendChild(WordServiceSetting.NormalParagraphWith_2Tabs("3. ฐานกฎหมายในการแบ่งปันข้อมูลส่วนบุคคล ภายใต้วัตถุประสงค์ที่ระบุในข้อ 2 คู่สัญญาแต่ละฝ่ายมีฐานกฎหมายตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลดังต่อไปนี้ ในการแบ่งปันข้อมูลส่วนบุคคลแก่คู่สัญญาอีกฝ่าย (แต่ละฝ่ายอาจใช้ฐานกฎหมายที่ต่างกันในการแบ่งปันข้อมูลส่วนบุคคล)", null, "32"));
                var legalBasisTable = new Table(
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
                            new TableCellProperties(new Shading { Fill = "FFFFFF" }),
                            WordServiceSetting.BoldParagraph("ฐานกฎหมายของ สสว.")
                        )
                    ),
                    new TableRow(
                        new TableCell(
                            WordServiceSetting.NormalParagraphWith_2TabsColor(
                                "1. (ระบุฐานกฎหมายในการแบ่งปันข้อมูลส่วนบุคคลของ สสว. เช่น เพื่อการให้บริการตามสัญญากับเจ้าของข้อมูลส่วนบุคคล)\n" +
                                "2. เพื่อการดำเนินการกิจสาธารณะหรือใช้อำนาจรัฐที่ สสว. ได้รับตาม...\n" +
                                "3. ได้รับความยินยอมในการเปิดเผยข้อมูลจากเจ้าของข้อมูลส่วนบุคคล",
                                null,
                                "FF0000"
                            )
                        )
                    ),
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new Shading { Fill = "FFFFFF" }),
                            WordServiceSetting.BoldParagraph("ฐานกฎหมายของ (ระบุชื่อคู่สัญญาอีกฝ่าย)")
                        )
                    ),
                    new TableRow(
                        new TableCell(
                            WordServiceSetting.NormalParagraphWith_2TabsColor(
                                "1. (ระบุฐานกฎหมายในการแบ่งปันข้อมูลส่วนบุคคลของคู่สัญญาอีกฝ่าย เช่น เพื่อการให้บริการตามสัญญากับเจ้าของข้อมูลส่วนบุคคล)\n" +
                                "2. ได้รับความยินยอมในการเปิดเผยข้อมูลจากเจ้าของข้อมูลส่วนบุคคล",
                                null,
                                "FF0000"
                            )
                        )
                    )
                );
                body.AppendChild(legalBasisTable);

                // ... (rest of the paragraphs, unchanged for brevity)
                // --- 6. Signature lines ---
                body.AppendChild(WordServiceSetting.EmptyParagraph());

                var signatureTable = new Table(
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
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.RightParagraph("ลงชื่อ................................................................."),
                            WordServiceSetting.CenteredParagraph("(............................................................)"),
                            WordServiceSetting.CenteredParagraph("ผู้อำนวยการสำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม หรือผู้ที่ผู้อำนวยการมอบหมาย")
                        ),
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.RightParagraph("ลงชื่อ................................................................."),
                            WordServiceSetting.CenteredParagraph("(............................................................)"),
                            WordServiceSetting.CenteredParagraph("............................................................")
                        )
                    ),
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.RightParagraph("ลงชื่อ................................................................."),
                            WordServiceSetting.CenteredParagraph("พยาน"),
                            WordServiceSetting.CenteredParagraph("(............................................................)"),
                            WordServiceSetting.CenteredParagraph("............................................................")
                        ),
                        new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50" }),
                            WordServiceSetting.RightParagraph("ลงชื่อ................................................................."),
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
        } catch (Exception ex) 
        
        { 
            throw new Exception("Error in WordEContract_DataPersonalService.OnGetWordContact_DataPersonalService: " + ex.Message);
        }
    
    }
    #endregion  4.1.1.2.6.บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล
}