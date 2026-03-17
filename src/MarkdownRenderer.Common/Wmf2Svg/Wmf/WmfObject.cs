using MarkdownRenderer.Wmf2Svg.Gdi;

namespace MarkdownRenderer.Wmf2Svg.Wmf;

public abstract class WmfObject : IGdiObject
{
    public int ID { get; set; }

    protected WmfObject(int id)
    {
        ID = id;
    }
}