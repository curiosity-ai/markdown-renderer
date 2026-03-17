using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Slides.Inlines;

public class EmphasisInlineRenderer : SlideObjectRenderer<EmphasisInline>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, EmphasisInline obj)
    {
        var runProps = new DocumentFormat.OpenXml.Drawing.RunProperties();
        if (renderer.TextFormat.TryPeek(out var currentProps))
        {
            runProps = (DocumentFormat.OpenXml.Drawing.RunProperties)currentProps.CloneNode(true);
        }

        if (obj.DelimiterCount >= 2)
        {
            runProps.Bold = true;
        }
        else
        {
            runProps.Italic = true;
        }

        renderer.TextFormat.Push(runProps);
        renderer.WriteChildren(obj);
        renderer.TextFormat.Pop();
    }
}
