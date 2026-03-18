using System;
using System.IO;

class Program
{
    static void Main()
    {
        var targetUri = new Uri("http://localhost/ppt/presentation.xml");
        var baseUri = new Uri("http://localhost/");
        Console.WriteLine(baseUri.MakeRelativeUri(targetUri));
    }
}
