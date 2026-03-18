using System.IO;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Markdig;
using Markdig.Syntax;
using Markdig.Renderers.Docx;
using Markdig.Renderers.Rtf;
using MarkdownRenderer.Writers;
using MarkdownRenderer.Primitives;
using System;
using System.Linq;
using D = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownRenderer.Markdown;

public class MarkdownConverter
{
    /// <summary>
    /// Gets or set the base absolute URI for resolving images which are specified as relative URIs.
    /// This can be a local folder path or an http(s) URL, allowing to process e.g. images stored on GitHub.
    /// If null or empty, only images which already use absolute URIs are processed.
    /// Once the absolute URI is retrieved, the converter will try to download online images or access local images.
    /// To prevent this, set the SkipImages property to true (false by default).
    /// Note: base64 images are not supported.
    /// </summary>
    public string? ImagesBaseUri { get; set; } = null;

    /// <summary>
    /// Gets or set the base absolute URI for resolving links which are specified as relative URIs.
    /// This can be a local folder path or an http(s) URL, allowing to keep e.g. relative GitHub links working.
    /// If null or empty, relative links will be kept as relative  
    /// (this can be desirable when rendering an offline Markdown file and saving the output document in the same folder).
    /// </summary>
    public string? LinksBaseUri { get; set; } = null;

    /// <summary>
    /// If set to true, the converter will not try to download online images or access local images.
    /// </summary>
    public bool SkipImages { get; set; } = false;

    /// <summary>
    /// If set to true, the converter will try to automatically download external resources (e.g. images) via HTTP.
    /// Disabled by default.
    /// </summary>
    public bool EnableExternalResources { get; set; } = false;

    /// <summary>
    /// Image converter to preserve WEBP and other image types when rendering Markdown. 
    /// If the MarkdownRenderer.ImageSharp or MarkdownRenderer.SystemDrawing package is installed, 
    /// this property can be set to a new instance of ImageSharpConverter or SystemDrawingConverter. 
    /// </summary>
    public IImageConverter? ImageConverter { get; set; } = null;

    /// <summary>
    /// Page size in millimetres (width, height). When set, applies to newly created DOCX/RTF documents.
    /// Values are expressed in millimetres; conversion to twips is performed internally.
    /// If null, the default template/page size is kept.
    /// </summary>
    /// <summary>
    /// Page size in millimetres. When set, applies to newly created DOCX/RTF documents.
    /// If null, the default template/page size is kept.
    /// </summary>
    public PageSize PageSize { get; set; } = PageSize.Default;

    /// <summary>
    /// Page margins in millimetres. When set, applies to newly created DOCX/RTF documents.
    /// If null, the default template/margins are kept.
    /// </summary>
    public PageMargins PageMargins { get; set; } = PageMargins.Default;

    /// <summary>
    /// Convert Markdown to PPTX slides.
    /// </summary>
    /// <param name="markdown">The input markdown source.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void ToSlides(MarkdownSource markdown, string outputFilePath)
    {
        using (var ms = new MemoryStream())
        {
            using (var document = PresentationDocument.Create(ms, PresentationDocumentType.Presentation))
            {
                InitializePresentation(document);
                var renderer = new Markdig.Renderers.Slides.SlideDocumentRenderer(document)
                {
                    ImagesBaseUri = this.ImagesBaseUri,
                    LinksBaseUri = this.LinksBaseUri,
                    SkipImages = this.SkipImages,
                    EnableExternalResources = this.EnableExternalResources
                };
                renderer.Render(markdown.Document);
                document.Save();
            }

            var fixedBytes = FixPresentationPackage(ms.ToArray());
            File.WriteAllBytes(outputFilePath, fixedBytes);
        }
    }

    /// <summary>
    /// Convert Markdown to PPTX slides.
    /// </summary>
    /// <param name="markdown">The input markdown source.</param>
    /// <param name="outputStream">The output stream.</param>
    public void ToSlides(MarkdownSource markdown, Stream outputStream)
    {
        using (var ms = new MemoryStream())
        {
            using (var document = PresentationDocument.Create(ms, PresentationDocumentType.Presentation))
            {
                InitializePresentation(document);
                var renderer = new Markdig.Renderers.Slides.SlideDocumentRenderer(document)
                {
                    ImagesBaseUri = this.ImagesBaseUri,
                    LinksBaseUri = this.LinksBaseUri,
                    SkipImages = this.SkipImages,
                    EnableExternalResources = this.EnableExternalResources
                };
                renderer.Render(markdown.Document);
                document.Save();
            }

            var fixedBytes = FixPresentationPackage(ms.ToArray());
            outputStream.Write(fixedBytes, 0, fixedBytes.Length);
        }
    }

    private byte[] FixPresentationPackage(byte[] packageBytes)
    {
        using (var inStream = new MemoryStream(packageBytes))
        using (var outStream = new MemoryStream())
        {
            using (var inArchive = new System.IO.Compression.ZipArchive(inStream, System.IO.Compression.ZipArchiveMode.Read))
            using (var outArchive = new System.IO.Compression.ZipArchive(outStream, System.IO.Compression.ZipArchiveMode.Create, true))
            {
                foreach (var entry in inArchive.Entries)
                {
                    var newEntry = outArchive.CreateEntry(entry.FullName);
                    using (var entryStream = entry.Open())
                    using (var newEntryStream = newEntry.Open())
                    {
                        if (entry.FullName == "[Content_Types].xml")
                        {
                            var doc = new System.Xml.XmlDocument();
                            doc.Load(entryStream);
                            var namespaceManager = new System.Xml.XmlNamespaceManager(doc.NameTable);
                            namespaceManager.AddNamespace("ct", "http://schemas.openxmlformats.org/package/2006/content-types");

                            var defaultXmlNode = doc.SelectSingleNode("//ct:Default[@Extension='xml']", namespaceManager) as System.Xml.XmlElement;
                            if (defaultXmlNode != null && defaultXmlNode.GetAttribute("ContentType") == "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")
                            {
                                defaultXmlNode.SetAttribute("ContentType", "application/xml");

                                var overrideNode = doc.CreateElement("Override", "http://schemas.openxmlformats.org/package/2006/content-types");
                                overrideNode.SetAttribute("PartName", "/ppt/presentation.xml");
                                overrideNode.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml");
                                doc.DocumentElement?.AppendChild(overrideNode);
                            }
                            doc.Save(newEntryStream);
                        }
                        else if (entry.FullName.EndsWith(".rels"))
                        {
                            var doc = new System.Xml.XmlDocument();
                            doc.Load(entryStream);
                            if (doc.DocumentElement != null)
                            {
                                foreach (System.Xml.XmlElement rel in doc.DocumentElement.ChildNodes.OfType<System.Xml.XmlElement>())
                                {
                                    if (rel.HasAttribute("TargetMode") && rel.GetAttribute("TargetMode") == "External") continue;
                                    var target = rel.GetAttribute("Target");
                                    if (target.StartsWith("/"))
                                    {
                                        if (entry.FullName == "_rels/.rels")
                                        {
                                            rel.SetAttribute("Target", target.Substring(1));
                                        }
                                        else
                                        {
                                            // Find the associated part name for the rels file
                                            // E.g., ppt/slides/_rels/slide1.xml.rels -> ppt/slides/slide1.xml
                                            var dirName = Path.GetDirectoryName(entry.FullName);
                                            var parentDir = Path.GetDirectoryName(dirName);
                                            var fileName = Path.GetFileNameWithoutExtension(entry.FullName);
                                            var partName = (parentDir != null && parentDir.Length > 0 ? parentDir + "/" : "") + fileName;
                                            partName = partName.Replace('\\', '/');

                                            var baseUri = new Uri("http://localhost/" + partName);
                                            var targetUri = new Uri("http://localhost" + target);
                                            var relUri = baseUri.MakeRelativeUri(targetUri).ToString();
                                            rel.SetAttribute("Target", relUri);
                                        }
                                    }
                                }
                            }
                            doc.Save(newEntryStream);
                        }
                        else
                        {
                            entryStream.CopyTo(newEntryStream);
                        }
                    }
                }
            }
            return outStream.ToArray();
        }
    }

    private void InitializePresentation(PresentationDocument presentationDoc)
    {
        // Core File Properties
        presentationDoc.PackageProperties.Creator = "MarkdownRenderer";
        presentationDoc.PackageProperties.Created = DateTime.UtcNow;
        presentationDoc.PackageProperties.Modified = DateTime.UtcNow;

        PresentationPart presentationPart = presentationDoc.AddPresentationPart();
        presentationPart.Presentation = new Presentation();

        // The correct OpenXML element order is: SlideMasterIdList, SlideIdList, SlideSize, NotesSize, DefaultTextStyle...
        var slideMasterIdList = new SlideMasterIdList();
        presentationPart.Presentation.AppendChild(slideMasterIdList);
        var slideIdList = new SlideIdList();
        presentationPart.Presentation.AppendChild(slideIdList);

        // Add SlideSize
        presentationPart.Presentation.AppendChild(new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 });
        presentationPart.Presentation.AppendChild(new NotesSize() { Cx = 6858000, Cy = 9144000 });

        // Add DefaultTextStyle
        presentationPart.Presentation.AppendChild(new DefaultTextStyle(
            new D.DefaultParagraphProperties(new D.DefaultRunProperties() { Language = "en-US" }),
            new D.Level1ParagraphProperties(new D.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 } ) { LeftMargin = 0, Alignment = D.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false },
            new D.Level2ParagraphProperties(new D.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 } ) { LeftMargin = 457200, Alignment = D.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false },
            new D.Level3ParagraphProperties(new D.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 } ) { LeftMargin = 914400, Alignment = D.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false },
            new D.Level4ParagraphProperties(new D.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 } ) { LeftMargin = 1371600, Alignment = D.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false },
            new D.Level5ParagraphProperties(new D.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 } ) { LeftMargin = 1828800, Alignment = D.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false },
            new D.Level6ParagraphProperties(new D.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 } ) { LeftMargin = 2286000, Alignment = D.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false },
            new D.Level7ParagraphProperties(new D.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 } ) { LeftMargin = 2743200, Alignment = D.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false },
            new D.Level8ParagraphProperties(new D.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 } ) { LeftMargin = 3200400, Alignment = D.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false },
            new D.Level9ParagraphProperties(new D.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 } ) { LeftMargin = 3657600, Alignment = D.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false }
        ));

        // Create Presentation Properties
        PresentationPropertiesPart presPropsPart = presentationPart.AddNewPart<PresentationPropertiesPart>();
        presPropsPart.PresentationProperties = new PresentationProperties(
            new PresentationPropertiesExtensionList(
                new PresentationPropertiesExtension() { Uri = "{E76CE94A-603C-4142-B9EB-6D1370010A27}" },
                new PresentationPropertiesExtension() { Uri = "{D31A062A-798A-4329-ABDD-BBA856620510}" },
                new PresentationPropertiesExtension() { Uri = "{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}" }
            )
        );

        // View Properties Part
        ViewPropertiesPart viewPropsPart = presentationPart.AddNewPart<ViewPropertiesPart>();
        viewPropsPart.ViewProperties = new ViewProperties(
            new NormalViewProperties(new RestoredLeft() { Size = 17989, AutoAdjust = false }, new RestoredTop() { Size = 94660 }),
            new SlideViewProperties(new CommonSlideViewProperties(new CommonViewProperties(new ScaleFactor(new D.ScaleX() { Numerator = 109, Denominator = 100 }, new D.ScaleY() { Numerator = 109, Denominator = 100 }), new Origin() { X = 854L, Y = 48L }), new GuideList()) { SnapToGrid = false }),
            new NotesTextViewProperties(new CommonViewProperties(new ScaleFactor(new D.ScaleX() { Numerator = 1, Denominator = 1 }, new D.ScaleY() { Numerator = 1, Denominator = 1 }), new Origin() { X = 0L, Y = 0L })),
            new GridSpacing() { Cx = 72008, Cy = 72008 }
        );

        // Table Styles Part
        TableStylesPart tableStylesPart = presentationPart.AddNewPart<TableStylesPart>();
        tableStylesPart.TableStyleList = new D.TableStyleList() { Default = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" };

        // Create Theme
        ThemePart themePart = presentationPart.AddNewPart<ThemePart>();
        themePart.Theme = new D.Theme() { Name = "Office Theme" };
        themePart.Theme.AppendChild(new D.ThemeElements(
            new D.ColorScheme(
                new D.Dark1Color(new D.SystemColor() { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
                new D.Light1Color(new D.SystemColor() { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
                new D.Dark2Color(new D.RgbColorModelHex() { Val = "44546A" }),
                new D.Light2Color(new D.RgbColorModelHex() { Val = "E7E6E6" }),
                new D.Accent1Color(new D.RgbColorModelHex() { Val = "5B9BD5" }),
                new D.Accent2Color(new D.RgbColorModelHex() { Val = "ED7D31" }),
                new D.Accent3Color(new D.RgbColorModelHex() { Val = "A5A5A5" }),
                new D.Accent4Color(new D.RgbColorModelHex() { Val = "FFC000" }),
                new D.Accent5Color(new D.RgbColorModelHex() { Val = "4472C4" }),
                new D.Accent6Color(new D.RgbColorModelHex() { Val = "70AD47" }),
                new D.Hyperlink(new D.RgbColorModelHex() { Val = "0563C1" }),
                new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "954F72" })
            ) { Name = "Office" },
            new D.FontScheme(
                new D.MajorFont(new D.LatinFont() { Typeface = "Calibri Light" }, new D.EastAsianFont() { Typeface = "" }, new D.ComplexScriptFont() { Typeface = "" }),
                new D.MinorFont(new D.LatinFont() { Typeface = "Calibri" }, new D.EastAsianFont() { Typeface = "" }, new D.ComplexScriptFont() { Typeface = "" })
            ) { Name = "Office" },
            new D.FormatScheme(
                new D.FillStyleList(
                    new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
                    new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
                    new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor })
                ),
                new D.LineStyleList(
                    new D.Outline(new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor })),
                    new D.Outline(new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor })),
                    new D.Outline(new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }))
                ),
                new D.EffectStyleList(
                    new D.EffectStyle(new D.EffectList()),
                    new D.EffectStyle(new D.EffectList()),
                    new D.EffectStyle(new D.EffectList())
                ),
                new D.BackgroundFillStyleList(
                    new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
                    new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
                    new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor })
                )
            ) { Name = "Office" }
        ));
        themePart.Theme.AppendChild(new D.ObjectDefaults());
        themePart.Theme.AppendChild(new D.ExtraColorSchemeList());

        // Create Slide Master
        SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(
                new Background(new BackgroundStyleReference(new D.SchemeColor() { Val = D.SchemeColorValues.Light1 }) { Index = 1001 }),
                new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties() { Id = 1, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new D.TransformGroup())
                )
            ),
            new ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
            new SlideLayoutIdList(),
            new TextStyles(
                new TitleStyle(new D.Level1ParagraphProperties(new D.DefaultRunProperties() { FontSize = 4400, Kerning = 1200 }) { Alignment = D.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false }),
                new BodyStyle(new D.Level1ParagraphProperties(new D.DefaultRunProperties() { FontSize = 2800, Kerning = 1200 }) { LeftMargin = 228600, Indent = -228600, Alignment = D.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false }),
                new OtherStyle(new D.Level1ParagraphProperties(new D.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }) { LeftMargin = 0, Alignment = D.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false })
            )
        );

        // Add Theme to Slide Master
        slideMasterPart.AddPart(themePart);

        // Create Slide Layout
        SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        slideLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(
                new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties() { Id = 1, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new D.TransformGroup()),
                    new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties() { Id = 2, Name = "Title 1" },
                            new NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })
                        ),
                        new ShapeProperties(
                            new D.Transform2D(new D.Offset() { X = 628650, Y = 365125 }, new D.Extents() { Cx = 7886700, Cy = 1325563 }),
                            new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle }
                        ),
                        new TextBody(new D.BodyProperties(), new D.ListStyle(), new D.Paragraph(new D.Run(new D.Text("Click to edit Master title style"))))
                    ),
                    new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties() { Id = 3, Name = "Text Placeholder 2" },
                            new NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Body, Index = 1 })
                        ),
                        new ShapeProperties(
                            new D.Transform2D(new D.Offset() { X = 628650, Y = 1825625 }, new D.Extents() { Cx = 7886700, Cy = 4351338 }),
                            new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle }
                        ),
                        new TextBody(new D.BodyProperties(), new D.ListStyle(), new D.Paragraph(new D.ParagraphProperties() { Level = 0 }, new D.Run(new D.Text("Click to edit Master text styles"))))
                    )
                )
            ),
            new ColorMapOverride(new D.MasterColorMapping())
        ) { Type = SlideLayoutValues.Text };

        // Link Slide Layout back to Slide Master
        slideMasterPart.SlideMaster.SlideLayoutIdList!.AppendChild(new SlideLayoutId() { Id = 2147483649, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) });

        // Add Blank Layout
        SlideLayoutPart blankLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        blankLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(
                new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties() { Id = 1, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new D.TransformGroup())
                )
            ),
            new ColorMapOverride(new D.MasterColorMapping())
        ) { Type = SlideLayoutValues.Blank };

        // Link Blank Layout to Slide Master
        slideMasterPart.SlideMaster.SlideLayoutIdList!.AppendChild(new SlideLayoutId() { Id = 2147483650, RelationshipId = slideMasterPart.GetIdOfPart(blankLayoutPart) });

        // Link Slide Master to Presentation
        presentationPart.Presentation.SlideMasterIdList!.AppendChild(new SlideMasterId() { Id = 2147483648, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) });
    }

    /// <summary>
    /// Convert Markdown to PPTX slides bytes.
    /// </summary>
    /// <param name="markdown">The input markdown source.</param>
    public byte[] ToSlidesBytes(MarkdownSource markdown)
    {
        using (var ms = new MemoryStream())
        {
            ToSlides(markdown, ms);
            return ms.ToArray();
        }
    }

    /// <summary>
    /// Convert Markdown to DOCX.
    /// </summary>
    /// <param name="markdown">The input markdown source.</param>
    /// <param name="outputStream">The output stream.</param>
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default). Has no effect when appending to an existing document.</param>
    /// <param name="append">If true, adds the converted Markdown to the DOCX content rather than replacing it (false by default). 
    /// The default styles are automatically inserted if the document does not override them.</param>
    public void ToDocx(MarkdownSource markdown, Stream outputStream, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document, bool append = false)
    {
        using (var document = ToWordprocessingDocument(markdown, outputStream, openXmlDocumentType, append))
        {
            document.Save();
        }
    }

    /// <summary>
    /// Convert Markdown to DOCX.
    /// </summary>
    /// <param name="markdown">The input markdown source.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default). Has no effect when appending to an existing document.</param>
    /// <param name="append">If true and the output file exists, adds the converted Markdown to the DOCX content rather than replacing it (false by default). 
    /// The default styles are automatically inserted if the document does not override them.</param>
    public void ToDocx(MarkdownSource markdown, string outputFilePath, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document, bool append = false)
    {
        using (var document = ToWordprocessingDocument(markdown, outputFilePath, openXmlDocumentType, append))
        {
            document.Save();
        }
    }

    /// <summary>
    /// Convert Markdown to DOCX bytes.
    /// </summary>
    /// <param name="markdown">The input markdown source.</param>
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default).</param>
    public byte[] ToDocxBytes(MarkdownSource markdown, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document)
    {
        using (var ms = new MemoryStream())
        {
            ToDocx(markdown, ms, openXmlDocumentType, false);
            return ms.ToArray();
        }
    }

    /// <summary>
    /// Convert Markdown to DOCX and returns a WordprocessingDocument for further modification. 
    /// The application is responsible for saving and disposing the document object.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="outputStream">The output stream.</param>
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default). Has no effect when appending to an existing document.</param>
    /// <param name="append">If true, adds the converted Markdown to the DOCX content rather than replacing it (false by default). 
    /// The default styles are automatically inserted if the document does not override them.</param>
    /// <returns>A <see cref="WordprocessingDocument"/> object.</returns>
    public WordprocessingDocument ToWordprocessingDocument(MarkdownSource markdown, Stream outputStream, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document, bool append = false)
    {
        WordprocessingDocument document;
        var defaultStyles = new DocumentStyles();
        if (append)
        {
            document = WordprocessingDocument.Open(outputStream, true);
            DocxTemplateHelper.AddStylesIfRequired(defaultStyles, document);
        }
        else
        {
            document = DocxTemplateHelper.BuildFromDefaultTemplate(outputStream, openXmlDocumentType);
        }
        var renderer = new DocxDocumentRenderer(document, defaultStyles)
        {
            ImagesBaseUri = this.ImagesBaseUri,
            LinksBaseUri = this.LinksBaseUri,
            SkipImages = this.SkipImages,
            EnableExternalResources = this.EnableExternalResources,
            ImageConverter = this.ImageConverter
        };
        renderer.Render(markdown.Document);
        // Apply page settings only for newly created documents (not when appending)
        if (!append)
        {
            ApplyPageSettingsToDocx(document);
        }
        return document;
    }

    /// <summary>
    /// Convert Markdown to DOCX and returns a WordprocessingDocument for further modification. 
    /// The application is responsible for saving and disposing the document object.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="outputFilePath">The output file path.</param>
    /// <param name="openXmlDocumentType">The Open XML document type (Document by default). Has no effect when appending to an existing document.</param>
    /// <param name="append">If true and the output file exists, adds the converted Markdown to the DOCX content rather than replacing it (false by default). 
    /// The default styles are automatically inserted if the document does not override them.</param>
    /// <returns>A <see cref="WordprocessingDocument"/> object.</returns>
    public WordprocessingDocument ToWordprocessingDocument(MarkdownSource markdown, string outputFilePath, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document, bool append = false)
    {
        WordprocessingDocument document;
        var defaultStyles = new DocumentStyles();
        if (append && File.Exists(outputFilePath))
        {
            document = WordprocessingDocument.Open(outputFilePath, true);
            DocxTemplateHelper.AddStylesIfRequired(defaultStyles, document);
        }
        else
        {
            document = DocxTemplateHelper.BuildFromDefaultTemplate(outputFilePath, openXmlDocumentType);
        }
        // Apply page settings only for newly created documents (not when appending)
        if (!append)
        {
            ApplyPageSettingsToDocx(document);
        }
        var renderer = new DocxDocumentRenderer(document, defaultStyles)
        {
            ImagesBaseUri = this.ImagesBaseUri,
            LinksBaseUri = this.LinksBaseUri,
            SkipImages = this.SkipImages,
            EnableExternalResources = this.EnableExternalResources,
            ImageConverter = this.ImageConverter
        };
        renderer.Render(markdown.Document);
        return document;
    }

    /// <summary>
    /// Convert Markdown to FlatOPC (flat XML-based DOCX variant) and returns an XDocument.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <returns>The FlatOPC as <see cref="XDocument"/></returns>
    public XDocument ToFlatOpc(MarkdownSource markdown)
    {
        using (var ms = new MemoryStream())
        {
            using (var document = ToWordprocessingDocument(markdown, ms))
            {
                document.Save();
                return document.ToFlatOpcDocument();
            }
        }
    }

    /// <summary>
    /// Convert Markdown to FlatOPC (flat XML-based DOCX variant) and returns a string.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <returns>The FlatOPC document as <see cref="string"/></returns>
    public string ToFlatOpcString(MarkdownSource markdown)
    {
        using (var ms = new MemoryStream())
        {
            using (var document = ToWordprocessingDocument(markdown, ms))
            {
                document.Save();
                return document.ToFlatOpcString();
            }
        }
    }

    /// <summary>
    /// Convert Markdown and append to an existing document instance.
    /// The application is responsible for saving and disposing the document.
    /// 
    /// The default styles required for Markdown conversion are automatically inserted,
    /// unless the document already contains a style with the same name 
    /// (this would take precedence over the defaults).
    /// </summary>
    /// <param name="markdown">The input markdown source.</param>
    /// <param name="outputDocument">The WordprocessingDocument to use.</param>
    public void AppendToDocument(MarkdownSource markdown, WordprocessingDocument outputDocument)
    {
        var styles = new DocumentStyles();
        DocxTemplateHelper.AddStylesIfRequired(styles, outputDocument);
        var renderer = new DocxDocumentRenderer(outputDocument, styles)
        {
            ImagesBaseUri = this.ImagesBaseUri,
            LinksBaseUri = this.LinksBaseUri,
            SkipImages = this.SkipImages,
            EnableExternalResources = this.EnableExternalResources,
            ImageConverter = this.ImageConverter
        };
        renderer.Render(markdown.Document);
    }

    /// <summary>
    /// Convert Markdown to RTF and save to file.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="outputFilePath">The output RTF file path.</param>
    public void ToRtf(MarkdownSource markdown, string outputFilePath, MarkdownToRtfSettings? settings = null)
    {
        using (var sw = new StreamWriter(outputFilePath, append: false, encoding: Encodings.UTF8NoBOM))
        {
            settings ??= new MarkdownToRtfSettings();
            var rtfBuilder = new RtfStringWriter() { ExternalWriter = sw };
            RenderToRtf(markdown.Document, rtfBuilder, settings);
        }
    }

    /// <summary>
    /// Convert Markdown to RTF and save to a Stream.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="outputStream">The output RTF stream.</param>
    public void ToRtf(MarkdownSource markdown, Stream outputStream, MarkdownToRtfSettings? settings = null)
    {
        using (var sw = new StreamWriter(outputStream, encoding: Encodings.UTF8NoBOM, bufferSize: 1024, leaveOpen: true))
        {
            settings ??= new MarkdownToRtfSettings();
            var rtfBuilder = new RtfStringWriter() { ExternalWriter = sw };
            RenderToRtf(markdown.Document, rtfBuilder, settings);
        }
    }

    /// <summary>
    /// Convert Markdown to RTF and write to a TextWriter.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="output">The output writer.</param>
    public void ToRtf(MarkdownSource markdown, TextWriter output, MarkdownToRtfSettings? settings = null)
    {
        settings ??= new MarkdownToRtfSettings();
        var rtfBuilder = new RtfStringWriter() { ExternalWriter = output };
        RenderToRtf(markdown.Document, rtfBuilder, settings);
    }

    /// <summary>
    /// Convert Markdown to RTF and returns a string.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <returns>The RTF document as <see cref="string"/></returns>
    public string ToRtfString(MarkdownSource markdown, MarkdownToRtfSettings? settings = null)
    {
        settings ??= new MarkdownToRtfSettings();
        var rtfBuilder = new RtfStringWriter();
        RenderToRtf(markdown.Document, rtfBuilder, settings);
        return rtfBuilder.ToString();
    }

    /// <summary>
    /// Convert Markdown to HTML and save to file.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="outputFilePath">The output HTML file path.</param>
    public void ToHtml(MarkdownSource markdown, string outputFilePath)
    {
        using (var sw = new StreamWriter(outputFilePath, append: false, encoding: Encodings.UTF8NoBOM))
            ToHtml(markdown, sw);
    }

    /// <summary>
    /// Convert Markdown to HTML and save to a Stream.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="outputStream">The output HTML stream.</param>
    public void ToHtml(MarkdownSource markdown, Stream outputStream)
    {
        using (var sw = new StreamWriter(outputStream, encoding: Encodings.UTF8NoBOM, bufferSize: 1024, leaveOpen: true))
            ToHtml(markdown, sw);
    }

    /// <summary>
    /// Convert Markdown to HTML and write to a TextWriter.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="output">The output writer.</param>
    public void ToHtml(MarkdownSource markdown, TextWriter output)
    {
        markdown.Document.ToHtml(output, MarkdownSource.DefaultPipeline);
        output.Flush();
    }

    /// <summary>
    /// Convert Markdown to HTML and returns a string.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <returns>The HTML document as <see cref="string"/></returns>
    public string ToHtmlString(MarkdownSource markdown)
    {
        return markdown.Document.ToHtml(MarkdownSource.DefaultPipeline);
    }

    /// <summary>
    /// Convert Markdown to plain text and save to file.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="outputFilePath">The output plain text (.txt) file path.</param>
    public void ToPlainText(MarkdownSource markdown, string outputFilePath)
    {
        using (var sw = new StreamWriter(outputFilePath, append: false, encoding: Encodings.UTF8NoBOM))
            ToPlainText(markdown, sw);
    }

    /// <summary>
    /// Convert Markdown to plain text and save to a Stream.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="outputStream">The output plain text stream.</param>
    public void ToPlainText(MarkdownSource markdown, Stream outputStream)
    {
        using (var sw = new StreamWriter(outputStream, encoding: Encodings.UTF8NoBOM, bufferSize: 1024, leaveOpen: true))
            ToPlainText(markdown, sw);
    }

    /// <summary>
    /// Convert Markdown to plain text and write to a TextWriter.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <param name="output">The output writer.</param>
    public void ToPlainText(MarkdownSource markdown, TextWriter output)
    {
        var renderer = new Markdig.Renderers.HtmlRenderer(output)
        {
            EnableHtmlForBlock = false,
            EnableHtmlForInline = false,
            EnableHtmlEscape = false,
        };
        MarkdownSource.DefaultPipeline.Setup(renderer);

        renderer.Render(markdown.Document);
        renderer.Writer.Flush(); // output.Flush();
    }

    /// <summary>
    /// Convert Markdown to plain text and returns a string.
    /// </summary>
    /// <param name="markdown">The markdown source.</param>
    /// <returns>The plain text as <see cref="string"/></returns>
    public string ToPlainTextString(MarkdownSource markdown)
    {
        using (var sw = new StringWriter())
        {
            ToPlainText(markdown, sw);
            return sw.ToString();
        }
    }
    
    private void RenderToRtf(MarkdownDocument document, RtfStringWriter rtfBuilder, MarkdownToRtfSettings settings)
    {
        var renderer = new RtfRenderer(rtfBuilder, settings)
        {
            ImagesBaseUri = this.ImagesBaseUri,
            LinksBaseUri = this.LinksBaseUri,
            ImageConverter = this.ImageConverter,
            SkipImages = this.SkipImages,
            EnableExternalResources = this.EnableExternalResources
        };
        // Apply page settings for RTF if configured
        if (PageSize != null)
        {
            renderer.PageWidthTwips = (int)PageSize.WidthTwips();
            renderer.PageHeightTwips = (int)PageSize.HeightTwips();
        }
        if (PageMargins != null)
        {
            renderer.MarginLeftTwips = (int)PageMargins.LeftTwips();
            renderer.MarginTopTwips = (int)PageMargins.TopTwips();
            renderer.MarginRightTwips = (int)PageMargins.RightTwips();
            renderer.MarginBottomTwips = (int)PageMargins.BottomTwips();
        }
        renderer.Render(document);
        rtfBuilder.Flush();
    }

    private void ApplyPageSettingsToDocx(WordprocessingDocument document)
    {
        if (document?.MainDocumentPart == null) return;
        document.MainDocumentPart.Document ??= new W.Document();
        document.MainDocumentPart.Document.Body ??= document.MainDocumentPart.Document.AppendChild(new W.Body());

        var sect = document.MainDocumentPart.Document.Body.Elements<W.SectionProperties>().LastOrDefault();
        if (sect == null)
        {
            sect = new W.SectionProperties();
            document.MainDocumentPart.Document.Body.AppendChild(sect);
        }

        // Page size
        if (PageSize != null)
        {
            uint width = (uint)PageSize.WidthTwips();
            uint height = (uint)PageSize.HeightTwips();
            var pgSize = sect.GetFirstChild<W.PageSize>() ?? new W.PageSize();
            pgSize.Width = width;
            pgSize.Height = height;
            if (sect.GetFirstChild<W.PageSize>() == null)
                sect.PrependChild(pgSize);
        }

        // Page margins
        if (PageMargins != null)
        {
            var pm = sect.GetFirstChild<W.PageMargin>() ?? new W.PageMargin();
            pm.Left = (uint)PageMargins.LeftTwips();
            pm.Right = (uint)PageMargins.RightTwips();
            pm.Top = (int)PageMargins.TopTwips();
            pm.Bottom = (int)PageMargins.BottomTwips();
            if (sect.GetFirstChild<W.PageMargin>() == null)
                sect.AppendChild(pm);
        }
    }
}
