using Markdig.Syntax;

namespace Markdig.Renderers.Slides.Blocks;

public class ListRenderer : SlideObjectRenderer<ListBlock>
{
    protected override void WriteObject(SlideDocumentRenderer renderer, ListBlock listBlock)
    {
        renderer.ActiveList.Push(new ListInfo
        {
            IsOrdered        = listBlock.IsOrdered,
            Start            = int.TryParse(listBlock.OrderedStart, out int start) ? start : 1,
            OrderedDelimiter = listBlock.OrderedDelimiter
        });
        renderer.WriteChildren(listBlock);
        renderer.ActiveList.Pop();
    }
}
