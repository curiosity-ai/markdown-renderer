using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Slides.Inlines;

public class LiteralInlineRenderer : SlideObjectRenderer<LiteralInline>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, LiteralInline obj)
    {
        WriteText(renderer, obj.Content.ToString());
    }
}
