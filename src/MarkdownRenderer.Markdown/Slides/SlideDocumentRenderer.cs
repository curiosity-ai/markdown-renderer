using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using MarkdownRenderer;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Markdig.Renderers.Slides.Blocks;
using Markdig.Renderers.Slides.Inlines;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Markdig.Renderers.Slides;

public class SlideDocumentRenderer : RendererBase
{
    public string? ImagesBaseUri { get; set; } = null;
    public string? LinksBaseUri { get; set; } = null;

    public bool SkipImages { get; set; } = false;

    public bool EnableExternalResources { get; set; } = false;

    public PresentationDocument Document { get; }

    public PresentationPart PresentationPart => Document.PresentationPart!;

    public SlidePart CurrentSlidePart { get; set; }

    public P.Shape CurrentShape { get; set; }
    public P.TextBody CurrentTextBody { get; set; }
    public D.Paragraph CurrentParagraph { get; set; }

    public Stack<D.RunProperties> TextFormat { get; } = new();

    // Some basic state tracking for lists
    public Stack<Markdig.Renderers.Slides.Blocks.ListInfo> ActiveList { get; } = new();

    private uint _slideIdCounter = 256;
    private uint _shapeIdCounter = 1000;

    public SlideDocumentRenderer(PresentationDocument document) : base()
    {
        Document = document;

        // Default block renderers
        ObjectRenderers.Add(new CodeBlockRenderer());
        ObjectRenderers.Add(new ListRenderer());
        ObjectRenderers.Add(new HeadingRenderer());
        ObjectRenderers.Add(new ParagraphRenderer());
        ObjectRenderers.Add(new ListItemRenderer());
        ObjectRenderers.Add(new QuoteBlockRenderer());
        ObjectRenderers.Add(new ThematicBreakRenderer());

        // Default inline renderers
        ObjectRenderers.Add(new AutolinkInlineRenderer());
        ObjectRenderers.Add(new CodeInlineRenderer());
        ObjectRenderers.Add(new DelimiterInlineRenderer());
        ObjectRenderers.Add(new EmphasisInlineRenderer());
        ObjectRenderers.Add(new LineBreakInlineRenderer());
        ObjectRenderers.Add(new HtmlEntityInlineRenderer());
        ObjectRenderers.Add(new LinkInlineRenderer());
        ObjectRenderers.Add(new LiteralInlineRenderer());
        ObjectRenderers.Add(new HtmlInlineRenderer());
    }

    public void CreateNewSlide()
    {
        var slidePart = PresentationPart.AddNewPart<SlidePart>();
        slidePart.Slide = new P.Slide(
            new P.CommonSlideData(
                new P.ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties() { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.GroupShapeProperties(new D.TransformGroup())
                )
            ),
            new P.ColorMapOverride(new D.MasterColorMapping())
        );

        var slideIdList = PresentationPart.Presentation.SlideIdList;
        if (slideIdList == null)
        {
            slideIdList = new P.SlideIdList();
            PresentationPart.Presentation.InsertAfter(slideIdList, PresentationPart.Presentation.SlideMasterIdList);
        }

        slideIdList.Append(new P.SlideId()
        {
            Id = _slideIdCounter++,
            RelationshipId = PresentationPart.GetIdOfPart(slidePart)
        });

        var slideLayoutPart = PresentationPart.SlideMasterParts.First().SlideLayoutParts.First();
        slidePart.AddPart(slideLayoutPart);

        slidePart.Slide.CommonSlideData!.ShapeTree = new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties() { Id = 1, Name = "" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new D.TransformGroup())
        );

        CurrentSlidePart = slidePart;
        CurrentShape = null!;
        CurrentTextBody = null!;
        CurrentParagraph = null!;
    }

    public P.Shape CreateShape(bool isTitle = false)
    {
        var shape = new P.Shape();

        if (isTitle)
        {
            shape.NonVisualShapeProperties = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = _shapeIdCounter++, Name = "Title" },
                new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape() { Type = P.PlaceholderValues.Title })
            );

            shape.ShapeProperties = new P.ShapeProperties(
                new D.Transform2D(
                    new D.Offset() { X = 457200L, Y = 274320L },
                    new D.Extents() { Cx = 8229600L, Cy = 1143000L }),
                new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle }
            );
        }
        else
        {
            shape.NonVisualShapeProperties = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = _shapeIdCounter++, Name = "Content Placeholder" },
                new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape() { Type = P.PlaceholderValues.Body, Index = 1 })
            );

            shape.ShapeProperties = new P.ShapeProperties(
                new D.Transform2D(
                    new D.Offset() { X = 457200L, Y = 1600200L },
                    new D.Extents() { Cx = 8229600L, Cy = 4525963L }),
                new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle }
            );
        }

        shape.TextBody = new P.TextBody(
            new D.BodyProperties(),
            new D.ListStyle()
        );

        return shape;
    }

    public D.Paragraph CreateParagraph()
    {
        // Content before the first level-1 heading has no slide yet.
        if (CurrentSlidePart == null)
        {
            CreateNewSlide();
        }

        var p = new D.Paragraph();
        if (ActiveList.Count > 0)
        {
            p.ParagraphProperties = CreateListParagraphProperties(ActiveList.Count, ActiveList.Peek());
        }

        if (CurrentTextBody == null) {
            CurrentShape = CreateShape();
            CurrentSlidePart.Slide.CommonSlideData!.ShapeTree!.AppendChild(CurrentShape);
            CurrentTextBody = CurrentShape.TextBody!;
        }
        CurrentTextBody.AppendChild(p);
        CurrentParagraph = p;
        return p;
    }

    private const int LIST_INDENT_EMU = 342900;

    private static D.ParagraphProperties CreateListParagraphProperties(int depth, Blocks.ListInfo listInfo)
    {
        var properties = new D.ParagraphProperties()
        {
            Level      = System.Math.Min(depth - 1, 8),
            LeftMargin = LIST_INDENT_EMU * depth,
            Indent     = listInfo.BulletPending ? -LIST_INDENT_EMU : 0
        };

        if (!listInfo.BulletPending)
        {
            properties.Append(new D.NoBullet());
        }
        else if (listInfo.IsOrdered)
        {
            properties.Append(new D.AutoNumberedBullet()
            {
                Type    = listInfo.OrderedDelimiter == ')' ? D.TextAutoNumberSchemeValues.ArabicParenR : D.TextAutoNumberSchemeValues.ArabicPeriod,
                StartAt = listInfo.Start + listInfo.Index - 1
            });
        }
        else
        {
            properties.Append(new D.CharacterBullet() { Char = depth % 2 == 1 ? "\u2022" : "\u2013" });
        }

        listInfo.BulletPending = false;
        return properties;
    }

    public override object Render(MarkdownObject markdownObject)
    {
        Write(markdownObject);
        return this;
    }
}
