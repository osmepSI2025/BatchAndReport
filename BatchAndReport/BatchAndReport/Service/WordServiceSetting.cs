﻿using BatchAndReport.Models;
using BatchAndReport.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

public class WordServiceSetting 
{
    public static Styles CreateDefaultStyles()
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


    public static Paragraph CenteredBoldParagraph(string text, string fontSize = "32") =>
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

    public static Paragraph CenteredParagraph(string text, string fontSize = "32") =>
        new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
            new Run(
                new RunProperties(
                    new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = fontSize } // Correct namespace and usage
                ),
                new Text(text)
            )
        );

    public static Paragraph RightParagraph(string text) =>
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

    public static Paragraph NormalParagraph(string text, JustificationValues? align = null, string fontSize = null) =>
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
    public static Paragraph EmptyParagraph() =>
        new Paragraph(new Run(
            new RunProperties(
                new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
                new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "32" } // Correct namespace and usage
            ),
            new Text("")
        ));

    public static Paragraph BoldUnderlineParagraph(string text) =>
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

    public static Paragraph BoldParagraph(string text) =>
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
    public static Drawing CreateImage(string relationshipId, long widthPx, long heightPx)
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

    public static Paragraph JustifiedParagraph_1tab(string text, string fontSize = "28", bool pitalic = false)
    {
        text = text.Replace(" ", "\u00A0");
        var runProps = new RunProperties(
            new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
            new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = fontSize }
        );
        if (pitalic)
            runProps.Append(new Italic());

        var props = new ParagraphProperties(new Justification { Val = JustificationValues.Both });
        props.Append(new Tabs(new TabStop { Val = TabStopValues.Left, Position = 720 }));

        return new Paragraph(
            props,
            new Run(runProps, new TabChar(), new Text(text))
        );
    }
    public static Paragraph JustifiedParagraph_2tab(string text, string fontSize = "28", bool pitalic = false)
    {
        text = text.Replace(" ", "\u00A0");
        var runProps = new RunProperties(
            new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
            new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = fontSize }
        );
        if (pitalic)
            runProps.Append(new Italic());

        var props = new ParagraphProperties(new Justification { Val = JustificationValues.Both });
        props.Append(new Tabs(new TabStop { Val = TabStopValues.Left, Position = 720 }));

        return new Paragraph(
            props,
            new Run(runProps, new TabChar(), new Text(text))
        );
    }
    public static Paragraph JustifiedParagraph(string text, string fontSize = "28", bool pitalic = false)
    {
        text = text.Replace(" ", "\u00A0");
        var runProps = new RunProperties(
            new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
            new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = fontSize }
        );
        if (pitalic)
            runProps.Append(new Italic());

        return new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Both }),

            new Run(runProps, new Text(text))
        );
    }
    // Helper: Paragraph with 2 tab spaces at the start of the first line
    public static Paragraph NormalParagraphWith_1Tabs(string text, JustificationValues? align = null, string fontZise = "28")
    {
        text = text.Replace(" ", "\u00A0");
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
    public static Paragraph NormalParagraphWith_2Tabs(string text, JustificationValues? align = null, string fontZise = "28", bool bold = false)
    {
      //  text = RemoveSpecialCharactersKeepingSomePunctuation(text);
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
            ,
            new TabStop { Val = TabStopValues.Left, Position = 1440 }
        );
        props.Append(tabs);
        paragraph.Append(props);

        // Add two tab characters at the start
        var runProps = new RunProperties(
            new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
            new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = fontZise }
        );
        if (bold)
            runProps.Append(new Bold());

        var run = new Run(
            runProps,
            new TabChar(),
            new TabChar(),
            new Text(text)
        );
        paragraph.Append(run);

        return paragraph;
    }
    public static Paragraph NormalParagraphWith_3Tabs(string text, JustificationValues? align = null, string fontZise = "28", bool bold = false)
    {
        if (fontZise == null)
        {
            fontZise = "28";
        }

        var paragraph = new Paragraph();

        // Paragraph properties: set justification to Both (full justify)
        var props = new ParagraphProperties(new Justification { Val = JustificationValues.Both });

        // Add three explicit tab stops for 0.5, 1.0, and 1.5 inches
        var tabs = new Tabs(
            new TabStop { Val = TabStopValues.Left, Position = 720 },
            new TabStop { Val = TabStopValues.Left, Position = 1440 },
            new TabStop { Val = TabStopValues.Left, Position = 2160 }
        );
        props.Append(tabs);
        paragraph.Append(props);

        // Correctly apply bold if requested
        var runProps = new RunProperties(
            new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
            new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = fontZise }
        );
        if (bold)
            runProps.Append(new Bold());

        var run = new Run(
            runProps,
            new TabChar(),
            new TabChar(),
            new TabChar(),
            new Text(text) { Space = SpaceProcessingModeValues.Preserve }
        );
        paragraph.Append(run);

        return paragraph;
    }
    //public static Paragraph NormalParagraphWith_3Tabs(string text, JustificationValues? align = null, string fontZise = "28", bool bold = false)
    //{
    //   // text = RemoveSpecialCharactersKeepingSomePunctuation(text);

    //    if (fontZise == null)
    //    {
    //        fontZise = "28";
    //    }

    //    var paragraph = new Paragraph();

    //    // Paragraph properties (alignment and tab stops)
    //    var props = new ParagraphProperties();
    //    if (align != null)
    //    {
    //        props.Append(new Justification { Val = align.Value });
    //    }


    //    // Add three explicit tab stops for 0.5, 1.0, and 1.5 inches
    //    var tabs = new Tabs(
    //        new TabStop { Val = TabStopValues.Left, Position = 720 },
    //        new TabStop { Val = TabStopValues.Left, Position = 1440 },
    //        new TabStop { Val = TabStopValues.Left, Position = 2160 }
    //    );
    //    props.Append(tabs);
    //    paragraph.Append(props);

    //    // Correctly apply bold if requested
    //    var runProps = new RunProperties(
    //        new RunFonts { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", EastAsia = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" },
    //        new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = fontZise }
    //    );
    //    if (bold)
    //        runProps.Append(new Bold());

    //    var run = new Run(
    //        runProps,
    //        new TabChar(),
    //        new TabChar(),
    //        new TabChar(),
    //        new Text(text) { Space = SpaceProcessingModeValues.Preserve }
    //    );
    //    paragraph.Append(run);

    //    return paragraph;
    //}

    public static Paragraph NormalParagraphWith_2TabsColor(string text, JustificationValues? align = null, string hexColor = null)
    {
        text = RemoveSpecialCharacters(text);
       // text = text.Replace(" ", "\u00A0");
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
       new Text(text) { Space = SpaceProcessingModeValues.Preserve }
        );
        paragraph.Append(run);

        return paragraph;
    }

    public static Paragraph CenteredBoldColoredParagraph(string text, string hexColor, string fonsize = "28") =>
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

    public static void AddHeaderWithPageNumber(MainDocumentPart mainPart, Body body)
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

    public static void AddHeaderWithLogoAndPageNumber(MainDocumentPart mainPart, Body body, string headerLogoRelId)
    {
        var headerPart = mainPart.AddNewPart<HeaderPart>();
        var header = new Header();

        // Add logo image to header
        var logoParagraph = new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
            CreateImage(headerLogoRelId, 120, 40)
        );
        header.Append(logoParagraph);

        // Add "hello" text to header (centered)
        var helloParagraph = CenteredBoldParagraph("hello", "32");
        header.Append(helloParagraph);

        // Add page number to header (right side)
        var pageNumParagraph = new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Right }),
            new Run(
                new FieldChar { FieldCharType = FieldCharValues.Begin },
                new FieldCode(" PAGE "),
                new FieldChar { FieldCharType = FieldCharValues.Separate },
                new Text("1"),
                new FieldChar { FieldCharType = FieldCharValues.End }
            )
        );
        header.Append(pageNumParagraph);

        headerPart.Header = header;

        // Reference header in section properties
        var sectionProps = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectionProps == null)
        {
            sectionProps = new SectionProperties();
            body.Append(sectionProps);
        }
        sectionProps.RemoveAllChildren<HeaderReference>();
        sectionProps.PrependChild(new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) });
    }

    public static string RemoveSpecialCharacters(string text)
    {
        // This regular expression matches any character that is NOT:
        // - an English letter (a-z, A-Z)
        // - a number (0-9)
        // - a whitespace character (\s)
        // - a Thai character (represented by the Unicode range \p{IsThai})
        // The '+' means match one or more occurrences.
        string cleanedText = Regex.Replace(text, @"[^a-zA-Z0-9\s\p{IsThai}]+", "");
        return cleanedText;
    }

    public static string RemoveSpecialCharactersKeepingSomePunctuation(string text)
    {
        // This version keeps English letters, numbers, spaces, Thai characters,
        // as well as periods (.), commas (,), and hyphens (-)
        string cleanedText = Regex.Replace(text, @"[^a-zA-Z0-9\s\p{IsThai}.,-]+", "");
        return cleanedText;
    }
}
