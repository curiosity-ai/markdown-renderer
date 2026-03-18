using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Markdig;
using Markdig.Extensions.Yaml;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace MarkdownRenderer.Markdown;

/// <summary>
/// Represents a source of Markdown content. It implicitly converts from a 
/// <see cref="Markdig.Syntax.MarkdownDocument"/> or <see cref="Stream" />,
/// or can be created from file path or <see cref="string"/> using  
/// the FromFile and FromMarkdownString static functions.
/// </summary>
public class MarkdownSource
{
    /// <summary>
    /// The Markdown Document associated to this source.
    /// </summary>
    public MarkdownDocument Document { get; }

    static MarkdownSource()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    /// <summary>
    /// Create a Markdown source from a document
    /// </summary>
    /// <param name="document">The <see cref="Markdig.Syntax.MarkdownDocument"/> to use</param>
    public MarkdownSource(MarkdownDocument document)
    {
        Document = document;
    }

    /// <summary>
    /// Create a Markdown source from a file path
    /// </summary>
    /// <param name="filePath">The file path to use</param>
    public static MarkdownSource FromFile(string filePath)
    {
        return FromMarkdownString(File.ReadAllText(filePath));
    }

    /// <summary>
    /// Create a Markdown source from a stream
    /// </summary>
    /// <param name="stream">The stream to use</param>
    /// <param name="encoding">The encoding to use (UTF8 by default if not specified).</param>
    public static MarkdownSource FromStream(Stream stream, Encoding? encoding = null)
    {
        using (var sr = new StreamReader(stream, encoding ?? Encoding.UTF8, detectEncodingFromByteOrderMarks: true, 
                                         bufferSize: 1024, leaveOpen: true))
        {
            string markdown = sr.ReadToEnd();
            return MarkdownSource.FromMarkdownString(markdown);
        }
    }

    internal static readonly MarkdownPipeline DefaultPipeline = new MarkdownPipelineBuilder()
                                                                        .UseAdvancedExtensions()
                                                                        .UseYamlFrontMatter()
                                                                        .UseEmojiAndSmiley()
                                                                        .Build();

    /// <summary>
    /// Create a Markdown source from a Markdown string
    /// </summary>
    /// <param name="markdown">The Markdown content as string</param>
    public static MarkdownSource FromMarkdownString(string markdown)
    {
        return Markdig.Markdown.Parse(markdown, DefaultPipeline);
    }

    /// <summary>
    /// Implicitly convert a <see cref="MarkdownDocument"/> to a <see cref="MarkdownSource"/>
    /// </summary>
    /// <param name="document">Markdown document to use</param>
    public static implicit operator MarkdownSource(MarkdownDocument document)
    {
        return new MarkdownSource(document);
    }

    /// <summary>
    /// Implicitly convert a <see cref="Stream"/> containing Markdown content to a <see cref="MarkdownSource"/>
    /// </summary>
    /// <param name="stream">Markdown content stream</param>
    public static implicit operator MarkdownSource(Stream stream)
    {
        return MarkdownSource.FromStream(stream);
    }

    /// <summary>
    /// Attempts to extract the title from the markdown input. It looks for a title field in the
    /// YAML frontmatter first, and if not found, it returns the text of the first header.
    /// </summary>
    /// <param name="markdown">The markdown input</param>
    /// <returns>The title, or null if no title was found</returns>
    public static string? TryGetTitle(string markdown)
    {
        if (string.IsNullOrWhiteSpace(markdown)) return null;

        var document = Markdig.Markdown.Parse(markdown, DefaultPipeline);

        foreach (var block in document)
        {
            if (block is YamlFrontMatterBlock yamlBlock)
            {
                var yaml = yamlBlock.Lines.ToString();
                var match = Regex.Match(yaml, @"^title:\s*(.*)", RegexOptions.IgnoreCase | RegexOptions.Multiline);
                if (match.Success)
                {
                    var title = match.Groups[1].Value.Trim().Trim('"', '\'');
                    if (!string.IsNullOrWhiteSpace(title))
                    {
                        return title;
                    }
                }
            }
            else if (block is HeadingBlock heading)
            {
                if (heading.Inline != null)
                {
                    var textBuilder = new StringBuilder();
                    foreach (var inline in heading.Inline)
                    {
                        if (inline is LiteralInline literal)
                        {
                            textBuilder.Append(literal.Content.ToString());
                        }
                        else if (inline is CodeInline code)
                        {
                            textBuilder.Append(code.Content);
                        }
                    }

                    var text = textBuilder.ToString();
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        return text;
                    }
                }
            }
        }

        return null;
    }
}

