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
    }
}
