using Markdig.Syntax;

namespace Markdig.Renderers.Slides.Blocks;

public abstract class ContainerBlockParagraphRendererBase<T> : ParagraphRendererBase<T> where T : ContainerBlock
{
    protected override void RenderContents(SlideDocumentRenderer renderer, T block)
    {
        renderer.WriteChildren(block);
    }
}
