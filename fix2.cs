using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;

class Program
{
    static void Main()
    {
        FixPackage("sample.pptx");
        Console.WriteLine("Done fixing.");
    }

    static void FixPackage(string path) {
        using (var package = Package.Open(path, FileMode.Open, FileAccess.ReadWrite))
        {
            var contentTypesPart = package.GetPart(new Uri("/[Content_Types].xml", UriKind.Relative));
            // Modify content types
            string contentTypesXml;
            using (var stream = contentTypesPart.GetStream(FileMode.Open, FileAccess.Read))
            using (var reader = new StreamReader(stream)) {
                contentTypesXml = reader.ReadToEnd();
            }
            contentTypesXml = contentTypesXml.Replace("<Default Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\" />",
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>");

            using (var stream = contentTypesPart.GetStream(FileMode.Create, FileAccess.Write))
            using (var writer = new StreamWriter(stream)) {
                writer.Write(contentTypesXml);
            }

            // Update root relationships target paths
            var rootRels = package.GetRelationships().ToList();
            foreach (var rel in rootRels) {
                if (rel.TargetMode == TargetMode.Internal) {
                    var uriStr = rel.TargetUri.ToString();
                    if (uriStr.StartsWith("/")) {
                        package.DeleteRelationship(rel.Id);
                        package.CreateRelationship(new Uri(uriStr.Substring(1), UriKind.Relative), TargetMode.Internal, rel.RelationshipType, rel.Id);
                    }
                }
            }

            // Update part relationships target paths
            foreach (var part in package.GetParts().ToList()) {
                var partRels = part.GetRelationships().ToList();
                foreach (var rel in partRels) {
                    if (rel.TargetMode == TargetMode.Internal) {
                        var uriStr = rel.TargetUri.ToString();
                        if (uriStr.StartsWith("/")) {
                            part.DeleteRelationship(rel.Id);
                            // compute relative Uri based on the current part URI
                            // e.g. from ppt/presentation.xml to ppt/slides/slide1.xml => slides/slide1.xml
                            var baseUri = new Uri("http://localhost" + part.Uri.ToString());
                            var targetUri = new Uri("http://localhost" + uriStr);
                            var relativeUri = baseUri.MakeRelativeUri(targetUri);
                            part.CreateRelationship(relativeUri, TargetMode.Internal, rel.RelationshipType, rel.Id);
                        }
                    }
                }
            }
        }
    }
}
