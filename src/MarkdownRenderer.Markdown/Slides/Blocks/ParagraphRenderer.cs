using Markdig.Syntax;

namespace Markdig.Renderers.Slides.Blocks;

public class ParagraphRenderer : LeafBlockParagraphRendererBase<ParagraphBlock>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, ParagraphBlock obj)
    {
        WriteAsParagraph(renderer, obj);
    }
}
