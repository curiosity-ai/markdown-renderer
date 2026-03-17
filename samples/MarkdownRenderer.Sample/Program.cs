using System;
using System.IO;
using MarkdownRenderer.Markdown;

namespace MarkdownRenderer.Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Generating Sample Documents from Markdown...");

            string markdown = @"# Slide 1

This is a paragraph on the first slide.
It has some **bold** and *italic* text.

* Item 1
* Item 2
  * Subitem A

# Slide 2

This is the second slide.

```csharp
// some code
int x = 42;
```

> A quote on the second slide.
";

            var converter = new MarkdownConverter();
            var source = MarkdownSource.FromMarkdownString(markdown);

            // Generate DOCX
            string docxPath = "sample.docx";
            converter.ToDocx(source, docxPath);
            Console.WriteLine($"Generated {docxPath}");

            // Generate PPTX
            string pptxPath = "sample.pptx";
            converter.ToSlides(source, pptxPath);
            Console.WriteLine($"Generated {pptxPath}");

            Console.WriteLine("Done!");
        }
    }
}
