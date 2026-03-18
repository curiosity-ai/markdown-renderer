using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Drawing;
using DocumentFormat.OpenXml.Packaging;
using Markdig.Syntax;
using MarkdownRenderer.Docx;
using MarkdownRenderer.IO;

namespace Markdig.Renderers.Docx.Blocks;

public class CodeBlockRenderer : LeafBlockParagraphRendererBase<CodeBlock>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, CodeBlock obj)
    {
        if (obj is FencedCodeBlock fcb && fcb.Info == "mermaid")
        {
            var diagram = obj.Lines.ToString();
            var base64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(diagram));
            var url = $"https://mermaid.ink/img/{base64}";

            try
            {
                using var http = new HttpClient();
                var bytes = http.GetByteArrayAsync(url).GetAwaiter().GetResult();
                using var stream = new MemoryStream(bytes);

                // Use the same width logic as full-page images or scale logic
                long widthInTwips = -1;
                long heightInTwips = -1;

                using (var tempStream = MarkdownRenderer.Markdown.Common.ImageHelpers.ConvertAndScaleImage(stream,
                                               out MarkdownRenderer.IO.ImageFormat fileType,
                                               renderer.Document.GetEffectivePageSize(),
                                               widthInTwips, heightInTwips,
                                               out long calculatedWidth, out long calculatedHeight,
                                               false, renderer.ImageConverter))
                {
                    PartTypeInfo? imagePartType = fileType.ToImagePartType();
                    if (tempStream != null && imagePartType != null && imagePartType.HasValue && calculatedWidth > 0 && calculatedHeight > 0)
                    {
                        tempStream.Position = 0;
                        var imagePart = renderer.Document.MainDocumentPart?.AddImagePart(tempStream, imagePartType.Value);
                        if (imagePart != null && renderer.Document.MainDocumentPart?.GetIdOfPart(imagePart) is string rId)
                        {
                            var imageElement = MarkdownRenderer.Docx.ImageHelpers.CreateImage(rId, calculatedWidth, calculatedHeight, (uint)(OpenXmlHelpers.GetImagesCount(renderer.Document) + 1), null, null);

                            var p = new Paragraph();
                            renderer.Cursor.Write(p);
                            renderer.Cursor.GoInto(p);
                            renderer.Cursor.Write(new Run(imageElement));
                            renderer.Cursor.PopAndAdvanceAfter(p);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                System.Diagnostics.Debug.WriteLine($"Exception in mermaid: {ex.Message}");
                #endif
                // Fallback to text rendering
                WriteAsParagraph(renderer, obj, renderer.Styles.MarkdownStyles["CodeBlock"]);
            }
        }
        else
        {
            WriteAsParagraph(renderer, obj, renderer.Styles.MarkdownStyles["CodeBlock"]);
        }
    }

    protected override void RenderContents(DocxDocumentRenderer renderer, CodeBlock obj)
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
