using System.Collections.Generic;

namespace MarkdownRenderer.Rtf;

internal class RtfDocument
{
    public RtfGroup Root { get; } = new RtfGroup();
}
