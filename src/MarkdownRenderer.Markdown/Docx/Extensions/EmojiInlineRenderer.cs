using Markdig.Extensions.Emoji;

namespace Markdig.Renderers.Docx.Extensions;

public class EmojiInlineRenderer : DocxObjectRenderer<EmojiInline>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, EmojiInline obj)
    {
        WriteText(renderer, obj.Content.ToString());
    }
}
