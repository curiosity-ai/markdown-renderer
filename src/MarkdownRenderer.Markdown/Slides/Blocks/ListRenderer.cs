using Markdig.Syntax;
using MarkdownRenderer.Docx;

namespace Markdig.Renderers.Slides.Blocks;

public class ListRenderer : SlideObjectRenderer<ListBlock>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, ListBlock listBlock)
    {
        // Simple list support
        renderer.ActiveList.Push(new ListInfo());
        renderer.WriteChildren(listBlock);
        renderer.ActiveList.Pop();
    }
}
