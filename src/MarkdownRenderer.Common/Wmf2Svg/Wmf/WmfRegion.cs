using MarkdownRenderer.Wmf2Svg.Gdi;

namespace MarkdownRenderer.Wmf2Svg.Wmf;

public class WmfRegion : WmfObject, IGdiRegion
{
    public WmfRegion(int id) : base(id)
    {
    }
}