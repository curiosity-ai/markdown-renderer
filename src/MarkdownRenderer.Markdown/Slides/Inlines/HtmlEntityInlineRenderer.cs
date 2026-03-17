using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Slides.Inlines;

public class HtmlEntityInlineRenderer : SlideObjectRenderer<HtmlEntityInline>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, HtmlEntityInline obj)
    {
        WriteText(renderer, obj.Transcoded.ToString());
    }
}
