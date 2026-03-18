using System;
using System.IO;
using DocumentFormat.OpenXml.Presentation;
using MarkdownRenderer.Markdown;

namespace MarkdownRenderer.Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Generating Sample Documents from Markdown...");

            string markdown = 
@"# Lists

This is a paragraph on the first slide.
It has some **bold** and *italic* text.

* Item 1
* Item 2
  * Subitem A

# Code Areas

This is the second slide.

```csharp
// some code
int x = 42;
```

> A quote on the second slide.

# Mermaid Diagrams

```mermaid
graph TD;
    A[Start] --> B{Is it?};
    B -- Yes --> C[OK];
    B -- No --> D[Cancel];
```


# Tables

| Header 1 | Header 2 |
| -------- | -------- |
| Cell 1   | Cell 2   |

# Mathematics

An inline equation $E=mc^2$ and a block equation:

$$
a^2 + b^2 = c^2
$$

# Custom Containers

:::callout
This is a custom container (callout)!.
:::

# Emoji and Smilies

Hello world! :smile: :rocket:
";




            string docxMarkdown = @"# Advanced Features for Word Document

## Tables

| Header 1 | Header 2 |
| -------- | -------- |
| Cell 1   | Cell 2   |

## Mathematics

An inline equation $E=mc^2$ and a block equation:

$$
a^2 + b^2 = c^2
$$

## Custom Containers

:::callout
This is a custom container (callout), which translates to a blockquote style in Word.
:::

## Emoji and Smilies

Hello world! :smile: :rocket:

## Footnotes

Here is a statement with a footnote.[^1]

[^1]: This is the text of the footnote.

## Mermaid Diagram

```mermaid
graph TD;
    A[Start] --> B{Is it?};
    B -- Yes --> C[OK];
    B -- No --> D[Cancel];
```
";

            var converter = new MarkdownConverter();
            var sourceDocx = MarkdownSource.FromMarkdownString(docxMarkdown);
            var sourceSlides = MarkdownSource.FromMarkdownString(markdown);

            // Generate DOCX
            string docxPath = "sample.docx";
            converter.ToDocx(sourceDocx, docxPath);
            Console.WriteLine($"Generated {docxPath}");

            // Generate PPTX
            string pptxPath = "sample.pptx";
            converter.ToSlides(sourceSlides, SlideSizeValues.Screen16x9, pptxPath);
            Console.WriteLine($"Generated {pptxPath}");

            Console.WriteLine("Done!");
        }
    }
}
