using Markdig.Syntax;

namespace Markdig.Renderers.Slides.Blocks;

public abstract class LeafBlockParagraphRendererBase<T> : ParagraphRendererBase<T> where T : LeafBlock
{
    protected override void RenderContents(SlideDocumentRenderer renderer, T block)
    {
        WriteLeafInline(renderer, block);
    }
}
