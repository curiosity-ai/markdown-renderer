using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Slides.Inlines;

public class LinkInlineRenderer : SlideObjectRenderer<LinkInline>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, LinkInline obj)
    {
        if (obj.IsImage)
        {
            // Placeholder for image handling
        }
        else
        {
            renderer.WriteChildren(obj);
        }
    }
}
