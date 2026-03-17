using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Slides.Inlines;

public class HtmlInlineRenderer : SlideObjectRenderer<HtmlInline>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, HtmlInline obj)
    {
        WriteText(renderer, obj.Tag);
    }
}
