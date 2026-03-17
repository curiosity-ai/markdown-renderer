using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MarkdownRenderer.Writers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownRenderer.Docx;

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    internal override void ProcessBookmarkStart(BookmarkStart bookmarkStart, RtfStringWriter sb)
    {
        sb.Write(@"{\*\bkmkstart " + bookmarkStart.Name + "}");
    }

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, RtfStringWriter sb)
    {
        sb.Write(@"{\*\bkmkend " + bookmarkEnd.GetBookmarkName() + "}");
    }
}
