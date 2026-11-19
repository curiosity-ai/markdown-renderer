using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;
using MarkdownRenderer.Docx;

namespace Markdig.Renderers.Docx.Blocks;

public class ListInfo
{
    public NumberingInstance? NumberingInstance { get; set; }
    
    public string? StyleId { get; set; }
    
    public int Level { get; set; }

    /// <summary>
    /// Left indentation of the item text at <see cref="Level"/> (from the numbering level), applied to the
    /// paragraphs of an item after its first one so they line up under the item.
    /// </summary>
    public string? ContinuationIndent { get; set; }
}

public class ListRenderer : DocxObjectRenderer<ListBlock>
{
    //private AbstractNum? _bulletListAbstractNum = null;
    //private AbstractNum? _orderedListAbstractNum = null;

    protected override void WriteObject(DocxDocumentRenderer renderer, ListBlock obj)
    {
        var listInfo = new ListInfo();
        var numbering = renderer.Document.GetOrCreateNumbering();
        var listItemStyle = obj.IsOrdered ? renderer.Styles.MarkdownStyles["ListOrderedItem"] : 
                                            renderer.Styles.MarkdownStyles["ListBulletItem"];
        listInfo.StyleId = listItemStyle;

        var abstractNumId = numbering == null ? null : FindAbstractNumId(renderer.Document, numbering, listItemStyle);
        if (abstractNumId != null)
        {
            var newNumberingId = numbering!.Elements<NumberingInstance>().Select(n => n.NumberID?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;
            var numberingInstance = new NumberingInstance
            {
                NumberID = newNumberingId,
                AbstractNumId = new AbstractNumId
                {
                    Val = abstractNumId.Value
                }
            };
            listInfo.NumberingInstance = numberingInstance;
            if (obj.IsOrdered)
            {
                for (var i = 0; i <= 8; i++)
                {
                    var lvlOverride = new LevelOverride
                    {
                        LevelIndex = i,
                        StartOverrideNumberingValue = new StartOverrideNumberingValue()
                        {
                            Val = int.TryParse(obj.OrderedStart, out int startNumber) ? startNumber : 1
                        }
                    };
                    listInfo.NumberingInstance.AppendChild(lvlOverride);
                }
            }
            numbering.AddNumberingInstance(listInfo.NumberingInstance);
        }
        // Without a numbering definition the items are still written, as plain paragraphs in the list style.

        if (renderer.ActiveList.Count == 0)
        {
            listInfo.Level = 0;
        }
        else
        {
            var previousList = renderer.ActiveList.Peek();
            listInfo.Level = Math.Min(previousList.Level + 1, 8); 
            // 8 seems to be the maximum level in DOCX documents (9 levels including 0).
        }
        if (abstractNumId != null)
        {
            var levelIndentation = numbering!.Elements<AbstractNum>()
                                             .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId)
                                             ?.Elements<Level>()
                                             .FirstOrDefault(l => l.LevelIndex?.Value == listInfo.Level)
                                             ?.PreviousParagraphProperties?.Indentation;
            listInfo.ContinuationIndent = levelIndentation?.Left?.Value ?? levelIndentation?.Start?.Value;
        }

        renderer.ActiveList.Push(listInfo);
        renderer.WriteChildren(obj);
        renderer.ActiveList.Pop();
    }

    /// <summary>
    /// Finds the abstract numbering definition behind a list paragraph style: the numbering instance named by the
    /// style's (or a base style's) w:numPr, which is how the default template and <c>AddStylesIfRequired</c> link them;
    /// then an abstract numbering that names the style through w:styleLink or a level's w:pStyle.
    /// </summary>
    private static int? FindAbstractNumId(WordprocessingDocument document, Numbering numbering, string styleId)
    {
        var styles = document.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (styles != null && currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
            if (style == null) break;

            var numId = style.StyleParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
            if (numId != null)
            {
                var abstractNumId = numbering.Elements<NumberingInstance>()
                                             .FirstOrDefault(n => n.NumberID?.Value == numId)
                                             ?.AbstractNumId?.Val?.Value;
                if (abstractNumId != null &&
                    numbering.Elements<AbstractNum>().Any(a => a.AbstractNumberId?.Value == abstractNumId))
                {
                    return abstractNumId;
                }
            }
            currentStyleId = style.BasedOn?.Val?.Value;
        }

        return numbering.Elements<AbstractNum>()
                        .FirstOrDefault(a => a.StyleLink?.Val?.Value == styleId ||
                                             a.Elements<Level>().Any(l => l.ParagraphStyleIdInLevel?.Val?.Value == styleId))
                        ?.AbstractNumberId?.Value;
    }
}
