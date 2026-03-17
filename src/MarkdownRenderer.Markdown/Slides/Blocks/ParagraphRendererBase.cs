using System;
using System.Linq;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using MarkdownRenderer.Docx;

namespace Markdig.Renderers.Slides.Blocks;

public abstract class ParagraphRendererBase<T> : SlideObjectRenderer<T> where T : Block
{
    protected DocumentFormat.OpenXml.Drawing.Paragraph WriteAsParagraph(SlideDocumentRenderer renderer, T obj)
    {
        var p = renderer.CreateParagraph();
        RenderContents(renderer, obj);
        return p;
    }

    protected abstract void RenderContents(SlideDocumentRenderer renderer, T block);
}
