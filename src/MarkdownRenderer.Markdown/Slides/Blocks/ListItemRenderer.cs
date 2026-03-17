using Markdig.Syntax;

namespace Markdig.Renderers.Slides.Blocks;

public class ListItemRenderer : ContainerBlockParagraphRendererBase<ListItemBlock>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, ListItemBlock obj)
    {
        RenderContents(renderer, obj);
    }
}
