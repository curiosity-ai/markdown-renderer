using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Slides.Inlines;

public class LineBreakInlineRenderer : SlideObjectRenderer<LineBreakInline>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, LineBreakInline obj)
    {
        if (obj.IsHard)
        {
            renderer.CreateParagraph();
        }
        else
        {
            WriteText(renderer, " ");
        }
    }
}
