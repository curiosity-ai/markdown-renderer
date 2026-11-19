using Markdig.Syntax;

namespace Markdig.Renderers.Slides.Blocks;

public class ListItemRenderer : ContainerBlockParagraphRendererBase<ListItemBlock>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, ListItemBlock obj)
    {
        var listInfo = renderer.ActiveList.Peek();
        listInfo.BulletPending = true;
        RenderContents(renderer, obj);
        listInfo.BulletPending = false;
        listInfo.Index++;
    }
}
