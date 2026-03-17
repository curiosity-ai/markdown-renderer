using Markdig.Syntax;

namespace Markdig.Renderers.Slides.Blocks;

public class ListItemRenderer : ContainerBlockParagraphRendererBase<ListItemBlock>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, ListItemBlock obj)
    {
        // Add a bullet paragraph
        var p = renderer.CreateParagraph();
        p.ParagraphProperties = new DocumentFormat.OpenXml.Drawing.ParagraphProperties()
        {
            Level = (renderer.ActiveList.Count > 0) ? renderer.ActiveList.Count : 1
        };
        RenderContents(renderer, obj);
    }
}
