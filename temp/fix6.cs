using System;
using System.IO;
using System.IO.Compression;
using System.Xml;

class Program
{
    static void Main()
    {
        FixPackage("../sample.pptx");
        Console.WriteLine("Done fixing.");
    }

    static void FixPackage(string path) {
        var tempFile = path + ".tmp";
        if (File.Exists(tempFile)) File.Delete(tempFile);

        using (var inArchive = ZipFile.Open(path, ZipArchiveMode.Read))
        using (var outArchive = ZipFile.Open(tempFile, ZipArchiveMode.Create))
        {
            foreach (var entry in inArchive.Entries)
            {
                var newEntry = outArchive.CreateEntry(entry.FullName);
                using (var stream = entry.Open())
                using (var outStream = newEntry.Open())
                {
                    if (entry.FullName == "[Content_Types].xml") {
                        using (var reader = new StreamReader(stream)) {
                            var contentTypesXml = reader.ReadToEnd();
                            contentTypesXml = contentTypesXml.Replace("<Default Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\" />",
                                "<Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>");
                            using (var writer = new StreamWriter(outStream)) {
                                writer.Write(contentTypesXml);
                            }
                        }
                    } else if (entry.FullName.EndsWith(".rels")) {
                        using (var reader = new StreamReader(stream)) {
                            var relsXml = reader.ReadToEnd();

                            var doc = new XmlDocument();
                            doc.LoadXml(relsXml);
                            foreach (XmlElement rel in doc.DocumentElement.ChildNodes) {
                                if (rel.HasAttribute("TargetMode") && rel.GetAttribute("TargetMode") == "External") continue;
                                var target = rel.GetAttribute("Target");
                                if (target.StartsWith("/")) {
                                    if (entry.FullName == "_rels/.rels") {
                                        rel.SetAttribute("Target", target.Substring(1));
                                    } else {
                                        var partName = entry.FullName.Replace("_rels/", "").Replace(".rels", "");
                                        var baseUri = new Uri("http://localhost/" + partName);
                                        var targetUri = new Uri("http://localhost" + target);
                                        var relUri = baseUri.MakeRelativeUri(targetUri).ToString();
                                        rel.SetAttribute("Target", relUri);
                                    }
                                }
                            }
                            using (var writer = new StreamWriter(outStream)) {
                                doc.Save(writer);
                            }
                        }
                    } else {
                        stream.CopyTo(outStream);
                    }
                }
            }
        }

        File.Delete(path);
        File.Move(tempFile, path);
    }
}
