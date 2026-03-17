using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Slides.Inlines;

public class DelimiterInlineRenderer : SlideObjectRenderer<DelimiterInline>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, DelimiterInline obj)
    {
        WriteText(renderer, obj.ToLiteral());
    }
}
