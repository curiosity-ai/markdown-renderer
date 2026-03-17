using Markdig.Syntax;

namespace Markdig.Renderers.Slides.Blocks;

public class CodeBlockRenderer : LeafBlockParagraphRendererBase<CodeBlock>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, CodeBlock obj)
    {
        var runProps = new DocumentFormat.OpenXml.Drawing.RunProperties();
        runProps.AppendChild(new DocumentFormat.OpenXml.Drawing.LatinFont() { Typeface = "Consolas" });
        renderer.TextFormat.Push(runProps);
        WriteAsParagraph(renderer, obj);
        renderer.TextFormat.Pop();
    }

    protected override void RenderContents(SlideDocumentRenderer renderer, CodeBlock obj)
    {
        if (obj is FencedCodeBlock)
        {
            var lines = obj.Lines.Lines;
            for (int i = 0; i < obj.Lines.Count; i++)
            {
                var line = lines[i];
                if (line.Slice.Text != null)
                {
                    WriteText(renderer, line.Slice.ToString());
                }

                if (i < obj.Lines.Count - 1)
                {
                    renderer.CreateParagraph();
                }
            }
        }
        else
        {
            base.RenderContents(renderer, obj);
        }
    }
}
