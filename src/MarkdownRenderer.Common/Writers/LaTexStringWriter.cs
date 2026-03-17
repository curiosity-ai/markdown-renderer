using System.Text;
using MarkdownRenderer.Helpers;

namespace MarkdownRenderer.Writers;

public sealed class LaTexStringWriter : BaseStringWriter
{
    public LaTexStringWriter()
    {
        NewLine = "\n"; // Use LF by default for LaTex
    }
}
