using Markdig.Extensions.CustomContainers;
using Markdig.Renderers.Docx.Blocks;
using MarkdownRenderer.Docx;

namespace Markdig.Renderers.Docx.Extensions;

public class CustomContainerRenderer : ContainerBlockParagraphRendererBase<CustomContainer>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, CustomContainer obj)
    {
        if (!renderer.IsFirstInContainer)
        {
            renderer.ForceCloseParagraph();
        }

        foreach (var block in obj)
        {
            var p = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            p.SetStyle(renderer.Styles.MarkdownStyles["Quote"]); // Map to styled callouts

            if (renderer.NoParagraph == 0)
            {
                renderer.Cursor.Write(p);
                renderer.Cursor.GoInto(p);
            }

            renderer.NoParagraph++;

            renderer.Write(block);

            if (renderer.NoParagraph == 0) continue;

            renderer.NoParagraph--;

            if (renderer.NoParagraph == 0)
            {
                renderer.Cursor.PopAndAdvanceAfter(p);
            }
        }
    }
}
