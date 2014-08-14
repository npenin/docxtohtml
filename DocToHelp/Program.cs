using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace DocToHelp
{
    class Program
    {
        static void Main(string[] args)
        {
            WordprocessingDocument doc = WordprocessingDocument.Open(args[0], false);
            if (Directory.Exists("Doc"))
                Directory.Delete("Doc", true);
            Directory.CreateDirectory("Doc");
            Environment.CurrentDirectory = Path.Combine(Environment.CurrentDirectory, "Doc");
            Directory.CreateDirectory("CSS");
            Directory.CreateDirectory("Images");
            Directory.CreateDirectory("Search");
            PageExtractor.Extract(doc.MainDocumentPart.Document.Body, doc);
            ImageExtractor.Extract(doc.MainDocumentPart.ImageParts);
            StyleExtractor.Extract(doc.MainDocumentPart.ThemePart, doc.MainDocumentPart.StyleDefinitionsPart);
            
            PageExtractor.ExtractToc();
        }
    }
}
