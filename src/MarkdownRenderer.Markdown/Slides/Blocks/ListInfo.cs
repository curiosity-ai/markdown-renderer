namespace Markdig.Renderers.Slides.Blocks;

public class ListInfo
{
    public int Index { get; set; } = 1;

    public bool IsOrdered { get; set; }

    public int Start { get; set; } = 1;

    public char OrderedDelimiter { get; set; } = '.';

    /// <summary>
    /// True while the current item has not written its first paragraph yet: that paragraph carries the bullet or
    /// number, the item's later paragraphs are indented continuations.
    /// </summary>
    public bool BulletPending { get; set; }
}
