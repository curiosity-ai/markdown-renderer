using System;
using System.Linq;
using System.Collections.Generic;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using MarkdownRenderer.Markdown;
using MarkdownRenderer.Markdown.Common;
using MarkdownRenderer.Helpers;
using MarkdownRenderer.Docx;
using BookmarkHelpers = MarkdownRenderer.Markdown.Common.BookmarkHelpers;
using D = DocumentFormat.OpenXml.Drawing;

namespace Markdig.Renderers.Slides.Blocks;

public class HeadingRenderer : LeafBlockParagraphRendererBase<HeadingBlock>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, HeadingBlock obj)
    {
        int level = Math.Min(Math.Max(obj.Level, 1), 6);

        if (level == 1)
        {
            renderer.CreateNewSlide();

            // Create title shape
            var shape = renderer.CreateShape(true);

            renderer.CurrentShape = shape;
            renderer.CurrentSlidePart.Slide.CommonSlideData!.ShapeTree!.AppendChild(shape);
            renderer.CurrentTextBody = shape.TextBody!;

            renderer.TextFormat.Push(new D.RunProperties() { FontSize = 4400, Bold = true });

            WriteAsParagraph(renderer, obj);

            renderer.TextFormat.Pop();

            // Reset to default shape so content blocks start a new body shape
            renderer.CurrentShape = null!;
            renderer.CurrentTextBody = null!;
            renderer.CurrentParagraph = null!;
        }
        else
        {
            renderer.TextFormat.Push(new D.RunProperties() { FontSize = 2800, Bold = true });
            WriteAsParagraph(renderer, obj);
            renderer.TextFormat.Pop();
        }
    }
}
