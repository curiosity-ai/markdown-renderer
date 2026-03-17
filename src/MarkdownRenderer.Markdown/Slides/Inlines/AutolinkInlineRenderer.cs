using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Slides.Inlines;

public class AutolinkInlineRenderer : SlideObjectRenderer<AutolinkInline>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, AutolinkInline obj)
    {
        WriteText(renderer, obj.Url);
    }
}
