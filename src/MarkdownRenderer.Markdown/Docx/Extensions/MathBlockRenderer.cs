using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Extensions.Mathematics;

namespace Markdig.Renderers.Docx.Extensions;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Renderers.Docx.Blocks;

public class MathBlockRenderer : LeafBlockParagraphRendererBase<MathBlock>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, MathBlock obj)
    {
        WriteAsParagraph(renderer, obj, renderer.Styles.MarkdownStyles["CodeBlock"]);
    }

    protected override void RenderContents(DocxDocumentRenderer renderer, MathBlock obj)
    {
        var lines = obj.Lines;

        for (var i = 0; i < lines.Count; i++)
        {
            var line = lines.Lines[i];
            var text = line.ToString() ?? "";

            var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            renderer.Cursor.Write(run);

            if (i < lines.Count - 1)
            {
                var breakRun = new Run(new Break());
                renderer.Cursor.Write(breakRun);
            }
        }
    }
}
