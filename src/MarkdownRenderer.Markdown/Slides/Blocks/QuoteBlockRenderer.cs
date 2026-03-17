using Markdig.Syntax;

namespace Markdig.Renderers.Slides.Blocks;

public class QuoteBlockRenderer : ContainerBlockParagraphRendererBase<QuoteBlock>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, QuoteBlock obj)
    {
        renderer.TextFormat.Push(new DocumentFormat.OpenXml.Drawing.RunProperties() { Italic = true });
        WriteAsParagraph(renderer, obj);
        renderer.TextFormat.Pop();
    }
}
