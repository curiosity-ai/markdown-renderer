using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Slides.Inlines;

public class CodeInlineRenderer : SlideObjectRenderer<CodeInline>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, CodeInline obj)
    {
        var runProps = new DocumentFormat.OpenXml.Drawing.RunProperties();
        if (renderer.TextFormat.TryPeek(out var currentProps))
        {
            runProps = (DocumentFormat.OpenXml.Drawing.RunProperties)currentProps.CloneNode(true);
        }

        runProps.AppendChild(new DocumentFormat.OpenXml.Drawing.LatinFont() { Typeface = "Consolas" });

        renderer.TextFormat.Push(runProps);
        WriteText(renderer, obj.Content);
        renderer.TextFormat.Pop();
    }
}
