using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToHelp
{
    class SearchExtractor
    {
        internal static void Extract(string textToIndex, string fileName, string heading)
        {
            foreach (string s in textToIndex.Split(new char[] { ':', '/', '\'', '\\', '"', '>', '*', ' ', '\'', '?', '!', '-', '(', ')', '.', ',', (char)0xA0 }, StringSplitOptions.RemoveEmptyEntries))
            {
                using (var index = new StreamWriter(File.Open("Search\\index-" + s + ".txt", FileMode.Append)))
                {
                    index.WriteLine(Path.GetFileName(fileName));
                    index.WriteLine(heading);
                    index.WriteLine(textToIndex);
                }
            }
        }
    }
}
