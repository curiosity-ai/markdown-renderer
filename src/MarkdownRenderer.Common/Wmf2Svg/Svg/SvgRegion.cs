using System.Xml;
using MarkdownRenderer.Wmf2Svg.Gdi;

namespace MarkdownRenderer.Wmf2Svg.Svg;

public abstract class SvgRegion : SvgObject, IGdiRegion
{
    protected SvgRegion(SvgGdi gdi) : base(gdi)
    {
    }

    public abstract XmlElement CreateElement();
}