using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using D = DocumentFormat.OpenXml.Drawing;
using MarkdownRenderer.Markdown;

namespace MarkdownRenderer.Markdown.Tests
{
    /// <summary>
    /// Regression tests: lists used to be dropped entirely from DOCX output because the list renderer
    /// looked the abstract numbering definition up by w:styleLink, which the default template does not use.
    /// </summary>
    public class ListRenderingTests
    {
        private sealed record ListItemInfo(string Text, int? NumId, int? Level, string? Format, int? Start);

        private static ListItemInfo[] RenderDocxListItems(string markdown, out string bodyText)
        {
            var converter = new MarkdownConverter();
            var bytes = converter.ToDocxBytes(MarkdownSource.FromMarkdownString(markdown));

            using var ms = new MemoryStream(bytes);
            using var doc = WordprocessingDocument.Open(ms, false);
            var main = doc.MainDocumentPart!;
            var numbering = main.NumberingDefinitionsPart?.Numbering;
            bodyText = main.Document!.Body!.InnerText;

            return main.Document.Body.Descendants<Paragraph>()
                       .Where(p => p.ParagraphProperties?.NumberingProperties != null)
                       .Select(p =>
                       {
                           var numPr = p.ParagraphProperties!.NumberingProperties!;
                           int? numId = numPr.NumberingId?.Val?.Value;
                           int? level = numPr.NumberingLevelReference?.Val?.Value;
                           string? format = null;
                           int? start = null;

                           var instance = numbering?.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID?.Value == numId);
                           var abstractId = instance?.AbstractNumId?.Val?.Value;
                           var abstractNum = numbering?.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == abstractId);
                           var lvl = abstractNum?.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == (level ?? 0));
                           format = lvl?.NumberingFormat?.Val?.ToString();
                           start = instance?.Elements<LevelOverride>()
                                            .FirstOrDefault(o => o.LevelIndex?.Value == (level ?? 0))
                                            ?.StartOverrideNumberingValue?.Val?.Value
                                   ?? lvl?.StartNumberingValue?.Val?.Value;

                           return new ListItemInfo(p.InnerText, numId, level, format, start);
                       })
                       .ToArray();
        }

        [Theory]
        [InlineData("- itemA\n- itemB\n", "bullet")]
        [InlineData("# T\n\n- itemA\n- itemB\n", "bullet")]
        [InlineData("* itemA\n* itemB\n", "bullet")]
        [InlineData("1. itemA\n2. itemB\n", "decimal")]
        public void Docx_SimpleList_HasNumberedItems(string markdown, string expectedFormat)
        {
            var items = RenderDocxListItems(markdown, out _);

            Assert.Equal(new[] { "itemA", "itemB" }, items.Select(i => i.Text));
            Assert.All(items, i =>
            {
                Assert.NotNull(i.NumId);
                Assert.Equal(0, i.Level);
                Assert.Equal(expectedFormat, i.Format);
            });
            Assert.Single(items.Select(i => i.NumId).Distinct());
        }

        [Fact]
        public void Docx_ListAfterParagraph_IsRendered()
        {
            var items = RenderDocxListItems("para\n\n- itemA\n", out var text);

            Assert.Contains("para", text);
            var item = Assert.Single(items);
            Assert.Equal("itemA", item.Text);
            Assert.Equal("bullet", item.Format);
        }

        [Fact]
        public void Docx_ListFollowedByQuote_RendersBoth()
        {
            var items = RenderDocxListItems("## H\n\n- itemA\n\n> q\n", out var text);

            Assert.Contains("H", text);
            Assert.Contains("q", text);
            var item = Assert.Single(items);
            Assert.Equal("itemA", item.Text);
            Assert.Equal("bullet", item.Format);
        }

        [Fact]
        public void Docx_SingleStarItem_IsRendered()
        {
            var item = Assert.Single(RenderDocxListItems("* itemA\n", out _));
            Assert.Equal("itemA", item.Text);
            Assert.Equal("bullet", item.Format);
        }

        [Fact]
        public void Docx_BulletLevels_UseUnicodeGlyphsWithoutSymbolFonts()
        {
            // Symbol / Wingdings bullets are private-use code points that have no glyph where those fonts are missing (Linux).
            var converter = new MarkdownConverter();
            var bytes = converter.ToDocxBytes(MarkdownSource.FromMarkdownString("- itemA\n"));

            using var ms = new MemoryStream(bytes);
            using var doc = WordprocessingDocument.Open(ms, false);
            var levels = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
                            .Elements<AbstractNum>()
                            .SelectMany(a => a.Elements<Level>())
                            .ToArray();
            Assert.NotEmpty(levels);

            foreach (var level in levels)
            {
                var text = level.LevelText?.Val?.Value ?? "";
                Assert.DoesNotContain(text, c => c >= '\uE000' && c <= '\uF8FF');

                var fonts = level.NumberingSymbolRunProperties?.RunFonts;
                var names = new[] { fonts?.Ascii?.Value, fonts?.HighAnsi?.Value, fonts?.ComplexScript?.Value, fonts?.EastAsia?.Value };
                Assert.DoesNotContain(names, n => n != null && (n.Contains("Symbol") || n.Contains("Wingdings")));
            }

            var bullet = levels.Where(l => l.NumberingFormat?.Val?.Value == NumberFormatValues.Bullet)
                               .OrderBy(l => l.LevelIndex!.Value)
                               .Take(3)
                               .Select(l => l.LevelText!.Val!.Value)
                               .ToArray();
            Assert.Equal(new[] { "\u2022", "\u25E6", "\u25AA" }, bullet);
        }

        [Fact]
        public void Docx_OrderedListStartingAtN_KeepsStart()
        {
            var items = RenderDocxListItems("3. three\n4. four\n", out _);

            Assert.Equal(new[] { "three", "four" }, items.Select(i => i.Text));
            Assert.All(items, i =>
            {
                Assert.Equal("decimal", i.Format);
                Assert.Equal(3, i.Start);
            });
        }

        [Fact]
        public void Docx_NestedLists_UseLevels()
        {
            var markdown = "- outer1\n  - inner1\n    1. deep1\n- outer2\n1. num1\n   - mixed\n";
            var items = RenderDocxListItems(markdown, out _);

            var byText = items.ToDictionary(i => i.Text);
            Assert.Equal(0, byText["outer1"].Level);
            Assert.Equal("bullet", byText["outer1"].Format);
            Assert.Equal(1, byText["inner1"].Level);
            Assert.Equal("bullet", byText["inner1"].Format);
            Assert.Equal(2, byText["deep1"].Level);
            Assert.Equal("lowerRoman", byText["deep1"].Format);
            Assert.Equal(0, byText["outer2"].Level);
            Assert.Equal(byText["outer1"].NumId, byText["outer2"].NumId);
            Assert.Equal(0, byText["num1"].Level);
            Assert.Equal("decimal", byText["num1"].Format);
            Assert.Equal(1, byText["mixed"].Level);
            Assert.Equal("bullet", byText["mixed"].Format);
        }

        [Fact]
        public void Docx_TaskList_RendersItems()
        {
            var items = RenderDocxListItems("- [ ] todo\n- [x] done\n", out _);

            Assert.Equal(2, items.Length);
            Assert.Contains("todo", items[0].Text);
            Assert.Contains("done", items[1].Text);
            Assert.All(items, i => Assert.Equal("bullet", i.Format));
        }

        [Fact]
        public void Docx_ListWithInlineFormatting_KeepsAllText()
        {
            var items = RenderDocxListItems("- **bold** and *em*\n- a [link](https://example.com)\n- some `code` here\n", out _);

            Assert.Equal(new[] { "bold and em", "a link", "some code here" }, items.Select(i => i.Text));
            Assert.All(items, i => Assert.Equal("bullet", i.Format));
        }

        [Fact]
        public void Docx_LooseListWithSeveralParagraphs_KeepsAllParagraphs()
        {
            var items = RenderDocxListItems("- first para\n\n  second para\n\n- next item\n\nafter\n", out var text);

            Assert.Equal(new[] { "first para", "next item" }, items.Select(i => i.Text));
            Assert.Contains("second para", text);
            Assert.Contains("after", text);
        }

        [Fact]
        public void Docx_LooseListContinuationParagraph_IsIndentedUnderItsItem()
        {
            var converter = new MarkdownConverter();
            var bytes = converter.ToDocxBytes(MarkdownSource.FromMarkdownString("- first para\n\n  second para\n\nafter\n"));

            using var ms = new MemoryStream(bytes);
            using var doc = WordprocessingDocument.Open(ms, false);
            var paragraphs = doc.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>().ToArray();

            var continuation = paragraphs.Single(p => p.InnerText == "second para");
            Assert.Null(continuation.ParagraphProperties?.NumberingProperties);
            Assert.Equal("720", continuation.ParagraphProperties?.Indentation?.Left?.Value);

            var after = paragraphs.Single(p => p.InnerText == "after");
            Assert.Null(after.ParagraphProperties?.Indentation);
        }

        [Fact]
        public void Docx_SeparateLists_RestartNumbering()
        {
            var items = RenderDocxListItems("1. a\n2. b\n\npara\n\n1. c\n", out _);

            Assert.Equal(3, items.Length);
            Assert.Equal(items[0].NumId, items[1].NumId);
            Assert.NotEqual(items[0].NumId, items[2].NumId);
        }

        [Theory]
        [InlineData("- itemA\n- itemB\n")]
        [InlineData("1. itemA\n2. itemB\n")]
        [InlineData("# T\n\n- itemA\n- itemB\n")]
        public void Slides_List_RendersItemText(string markdown)
        {
            var converter = new MarkdownConverter();
            var bytes = converter.ToSlidesBytes(MarkdownSource.FromMarkdownString(markdown), SlideSizeValues.Screen16x9);

            using var ms = new MemoryStream(bytes);
            using var pres = PresentationDocument.Open(ms, false);
            var text = string.Join("\n", pres.PresentationPart!.SlideParts.Select(s => s.Slide!.InnerText));

            Assert.Contains("itemA", text);
            Assert.Contains("itemB", text);
        }
        private static D.Paragraph[] RenderSlideParagraphs(string markdown)
        {
            var converter = new MarkdownConverter();
            var bytes = converter.ToSlidesBytes(MarkdownSource.FromMarkdownString(markdown), SlideSizeValues.Screen16x9);

            using var ms = new MemoryStream(bytes);
            using var pres = PresentationDocument.Open(ms, false);
            return pres.PresentationPart!.SlideParts
                       .SelectMany(s => s.Slide!.Descendants<D.Paragraph>())
                       .Select(p => (D.Paragraph)p.CloneNode(true))
                       .ToArray();
        }

        [Fact]
        public void Slides_Lists_HaveBulletsNumbersAndLevels()
        {
            var paragraphs = RenderSlideParagraphs("# T\n\n- outer\n  - inner\n\n3. three\n4. four\n\n- first\n\n  more\n");
            D.Paragraph ByText(string text) => paragraphs.Single(p => p.InnerText == text);

            var outer = ByText("outer").ParagraphProperties!;
            Assert.Equal(0, outer.Level!.Value);
            Assert.NotNull(outer.GetFirstChild<D.CharacterBullet>());

            var inner = ByText("inner").ParagraphProperties!;
            Assert.Equal(1, inner.Level!.Value);
            Assert.NotNull(inner.GetFirstChild<D.CharacterBullet>());

            Assert.Equal(3, ByText("three").ParagraphProperties!.GetFirstChild<D.AutoNumberedBullet>()!.StartAt!.Value);
            Assert.Equal(4, ByText("four").ParagraphProperties!.GetFirstChild<D.AutoNumberedBullet>()!.StartAt!.Value);

            Assert.NotNull(ByText("first").ParagraphProperties!.GetFirstChild<D.CharacterBullet>());
            Assert.NotNull(ByText("more").ParagraphProperties!.GetFirstChild<D.NoBullet>());
        }
    }
}
