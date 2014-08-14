using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace DocToHelp
{
    class ImageExtractor
    {
        internal static void Extract(IEnumerable<DocumentFormat.OpenXml.Packaging.ImagePart> images)
        {
            byte[] buffer = new byte[1024];
            int length;
            foreach (var image in images)
            {
                string imageName = image.Uri.ToString();
                imageName = imageName.Substring(imageName.LastIndexOf('/') + 1);
                using (var file = File.OpenWrite("Images\\" + imageName))
                {
                    var stream = image.GetStream();
                    while ((length = stream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        file.Write(buffer, 0, length);
                    }
                    file.Close();
                }
            }
        }
    }
}
