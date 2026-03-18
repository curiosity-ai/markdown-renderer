using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace MarkdownRenderer.Tools
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: MarkdownRenderer.Tools <file1.pptx|docx> <file2.pptx|docx>");
                return;
            }

            var file1 = args[0];
            var file2 = args[1];

            if (!File.Exists(file1)) { Console.WriteLine($"File not found: {file1}"); return; }
            if (!File.Exists(file2)) { Console.WriteLine($"File not found: {file2}"); return; }

            using (var archive1 = ZipFile.OpenRead(file1))
            using (var archive2 = ZipFile.OpenRead(file2))
            {
                var entries1 = archive1.Entries.Select(e => e.FullName).ToList();
                var entries2 = archive2.Entries.Select(e => e.FullName).ToList();

                var onlyIn1 = entries1.Except(entries2).ToList();
                var onlyIn2 = entries2.Except(entries1).ToList();

                if (onlyIn1.Any())
                {
                    Console.WriteLine($"\nFiles only in {file1}:");
                    foreach (var file in onlyIn1) Console.WriteLine($"  {file}");
                }

                if (onlyIn2.Any())
                {
                    Console.WriteLine($"\nFiles only in {file2}:");
                    foreach (var file in onlyIn2) Console.WriteLine($"  {file}");
                }

                var common = entries1.Intersect(entries2).ToList();
                foreach (var file in common)
                {
                    var entry1 = archive1.GetEntry(file);
                    var entry2 = archive2.GetEntry(file);

                    if (entry1 == null || entry2 == null) continue;

                    if (entry1.Length != entry2.Length)
                    {
                        if (file.EndsWith(".xml") || file.EndsWith(".rels"))
                        {
                            using (var stream1 = entry1.Open())
                            using (var stream2 = entry2.Open())
                            {
                                try
                                {
                                    var doc1 = XDocument.Load(stream1);
                                    var doc2 = XDocument.Load(stream2);

                                    var str1 = doc1.ToString(SaveOptions.DisableFormatting);
                                    var str2 = doc2.ToString(SaveOptions.DisableFormatting);

                                    if (str1 != str2)
                                    {
                                       Console.WriteLine($"\nFiles differ: {file} ({entry1.Length} vs {entry2.Length} bytes)");
                                       Console.WriteLine($"   XML structures differ.");
                                    }
                                }
                                catch
                                {
                                    Console.WriteLine($"\nFiles differ: {file} ({entry1.Length} vs {entry2.Length} bytes)");
                                    Console.WriteLine($"   Failed to parse as XML.");
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine($"\nFiles differ: {file} ({entry1.Length} vs {entry2.Length} bytes)");
                        }
                    }
                }
            }
        }
    }
}
