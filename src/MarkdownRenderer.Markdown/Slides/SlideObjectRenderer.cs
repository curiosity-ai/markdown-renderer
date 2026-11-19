using System;
using DocumentFormat.OpenXml;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using MarkdownRenderer.Helpers;
using D = DocumentFormat.OpenXml.Drawing;

namespace Markdig.Renderers.Slides;

public abstract class SlideObjectRenderer<T> : MarkdownObjectRenderer<SlideDocumentRenderer, T> where T : MarkdownObject
{
    protected override void Write(SlideDocumentRenderer renderer, T obj)
    {
        if (renderer == null) throw new ArgumentNullException(nameof(renderer));
        if (obj == null) throw new ArgumentNullException(nameof(obj));

        WriteObject(renderer, obj);
    }

    public void WriteText(SlideDocumentRenderer renderer, string text)
    {
        var run = new D.Run(new D.Text(text));

        if (renderer.TextFormat.TryPeek(out var props))
        {
            run.RunProperties = new D.RunProperties(props!.OuterXml);
        }
        else
        {
            run.RunProperties = new D.RunProperties() { Language = "en-US", FontSize = 1800, Dirty = false };
        }

        if (renderer.CurrentParagraph == null)
        {
            renderer.CreateParagraph();
        }

        renderer.CurrentParagraph.AppendChild(run);
    }

    public void WriteLeafInline(SlideDocumentRenderer renderer, LeafBlock leafBlock)
    {
        if (leafBlock is null || leafBlock.Inline is null)
            return;

        var inline = (Inline)leafBlock.Inline;
        while (inline != null)
        {
            renderer.Write(inline);
            inline = inline.NextSibling;
        }
    }

    protected abstract void WriteObject(SlideDocumentRenderer renderer, T obj);
}
