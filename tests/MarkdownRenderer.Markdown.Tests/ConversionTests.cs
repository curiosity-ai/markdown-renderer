using System.IO;
using Xunit;
using MarkdownRenderer.Markdown;

namespace MarkdownRenderer.Markdown.Tests
{
    public class ConversionTests
    {
        private readonly string markdownSample = @"
# Slide 1

This is a paragraph with **bold** and *italic* text.

* Item 1
* Item 2
  * Subitem

# Slide 2

This is the second slide.

```csharp
int x = 42;
```

> A quote.

---

End of slide
";

        [Fact]
        public void ToDocx_DoesNotThrow()
        {
            var converter = new MarkdownConverter();
            var source = MarkdownSource.FromMarkdownString(markdownSample);

            using var ms = new MemoryStream();

            // Act & Assert
            var ex = Record.Exception(() => converter.ToDocx(source, ms));
            Assert.Null(ex);
            Assert.True(ms.Length > 0);
        }

        [Fact]
        public void ToRtf_DoesNotThrow()
        {
            var converter = new MarkdownConverter();
            var source = MarkdownSource.FromMarkdownString(markdownSample);

            using var ms = new MemoryStream();

            // Act & Assert
            var ex = Record.Exception(() => converter.ToRtf(source, ms));
            Assert.Null(ex);
            Assert.True(ms.Length > 0);
        }

        [Fact]
        public void ToSlides_DoesNotThrow()
        {
            var converter = new MarkdownConverter();
            var source = MarkdownSource.FromMarkdownString(markdownSample);

            using var ms = new MemoryStream();

            // Act & Assert
            var ex = Record.Exception(() => converter.ToSlides(source, ms));
            Assert.Null(ex);
            Assert.True(ms.Length > 0);
        }

        [Fact]
        public void ToSlidesBytes_DoesNotThrow()
        {
            var converter = new MarkdownConverter();
            var source = MarkdownSource.FromMarkdownString(markdownSample);

            // Act & Assert
            var bytes = converter.ToSlidesBytes(source);
            Assert.NotNull(bytes);
            Assert.True(bytes.Length > 0);
        }

        [Fact]
        public void TryGetTitle_WithYamlFrontMatter_ReturnsTitle()
        {
            var markdown = "---\n" +
                           "title: \"My YAML Title\"\n" +
                           "---\n" +
                           "# Main Heading\n" +
                           "Some text.";
            var title = MarkdownSource.TryGetTitle(markdown);
            Assert.Equal("My YAML Title", title);
        }

        [Fact]
        public void TryGetTitle_WithHeading_ReturnsFirstHeading()
        {
            var markdown = "Some text.\n" +
                           "# My First Heading\n" +
                           "## Subheading";
            var title = MarkdownSource.TryGetTitle(markdown);
            Assert.Equal("My First Heading", title);
        }

        [Fact]
        public void TryGetTitle_WithCodeInHeading_ReturnsText()
        {
            var markdown = "# Header with `code` block\n";
            var title = MarkdownSource.TryGetTitle(markdown);
            Assert.Equal("Header with code block", title);
        }

        [Fact]
        public void TryGetTitle_WithNoTitle_ReturnsNull()
        {
            var markdown = "Just some regular text.\n" +
                           "No headings or frontmatter here.";
            var title = MarkdownSource.TryGetTitle(markdown);
            Assert.Null(title);
        }

        [Fact]
        public void TryGetTitle_WithYamlAndHeading_ReturnsYamlTitle()
        {
            var markdown = "---\n" +
                           "title: 'YAML Title'\n" +
                           "---\n" +
                           "# Markdown Heading\n";
            var title = MarkdownSource.TryGetTitle(markdown);
            Assert.Equal("YAML Title", title);
        }
    }
}
