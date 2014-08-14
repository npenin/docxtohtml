using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;
using System.Xml;
using DocumentFormat.OpenXml.Vml;
using System.Text.RegularExpressions;

namespace DocToHelp
{
    class PageExtractor
    {
        internal static void Extract(Body body, WordprocessingDocument doc)
        {
            XmlDocument tocDoc = new XmlDocument();
            XmlElement currentElement = tocDoc.CreateElement("toc");
            var levelAtt = tocDoc.CreateAttribute("level");
            currentElement.Attributes.Append(levelAtt);
            levelAtt.Value = "0";
            string currentHeading = "";
            tocDoc.AppendChild(currentElement);
            TextWriter writer = CreateTextFile("page", true);
            bool ulStarted = false;

            foreach (OpenXmlElement oxe in body)
            {
                SdtBlock toc = oxe as SdtBlock;
                if (toc != null)
                {
                    foreach (OpenXmlElement oxe2 in toc.SdtContentBlock)
                    {
                        var sdt = oxe2 as SdtBlock;
                        if (sdt == null)
                            continue;
                        XmlDocument xmldoc = new XmlDocument();
                        var xmlns = new XmlNamespaceManager(xmldoc.NameTable);
                        xmlns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                        xmldoc.LoadXml(sdt.SdtContentBlock.OuterXml);
                        foreach (XmlNode node in xmldoc.SelectNodes("//w:p[w:hyperlink]", xmlns))
                        {
                            var tocElement = tocDoc.CreateElement("level");
                            var level = new Regex("^[A-Z]+").Replace(node.SelectSingleNode("w:pPr/w:pStyle/@w:val", xmlns).Value, "");
                            
                            for (int i = int.Parse(currentElement.Attributes["level"].Value) - int.Parse(level); i >= 0; i--)
                                currentElement = (XmlElement)currentElement.ParentNode;
                            currentElement.AppendChild(tocElement);
                            currentElement = tocElement;
                            levelAtt = tocDoc.CreateAttribute("level");
                            tocElement.Attributes.Append(levelAtt);

                            levelAtt.Value = level;
                            StringBuilder text = new StringBuilder();
                            foreach (XmlNode t in node.SelectNodes("w:hyperlink/w:r[not(w:rPr/w:webHidden)]/w:t/text()", xmlns))
                                text.Append(t.Value);
                            var textAttr = tocDoc.CreateAttribute("text");
                            tocElement.Attributes.Append(textAttr);
                            textAttr.Value = text.ToString();
                        }
                    }
                    using (var tocxml = File.OpenWrite("toc.xml"))
                        tocDoc.Save(new StreamWriter(tocxml, Encoding.UTF8));
                }

                Paragraph p = oxe as Paragraph;
                if (p == null)
                    continue;

                if (p.ChildElements.OfType<BookmarkStart>().Any() && p.InnerText.Length > 0)
                {
                    var properties = p.GetFirstChild<ParagraphProperties>();

                    if (properties != null
                        && properties.ParagraphStyleId != null
                        && properties.ParagraphStyleId.Val != null
                        && properties.ParagraphStyleId.Val.HasValue
                        && (properties.ParagraphStyleId.Val.Value.StartsWith("Heading") || properties.ParagraphStyleId.Val.Value.StartsWith("Titre")))
                    {
                        writer.Close();
                        currentHeading = p.InnerText;
                        writer = CreateTextFile(currentHeading, true);
                    }
                }


                string startTag = "<p";
                if (p.ParagraphProperties != null && p.ParagraphProperties.NumberingProperties != null && (p.ParagraphProperties.ParagraphStyleId == null || !p.ParagraphProperties.ParagraphStyleId.Val.Value.ToLower().StartsWith("heading")))
                {
                    if (!ulStarted)
                    {
                        writer.Write("<ul>");
                        ulStarted = true;
                    }
                    startTag = "<li";
                }
                else if (ulStarted)
                {
                    writer.Write("</ul>");
                    ulStarted = false;
                }

                writer.Write(startTag);
                if (p.ParagraphProperties != null)
                    foreach (var oxeProperties in p.ParagraphProperties)
                    {
                        ParagraphStyleId pStyle = oxeProperties as ParagraphStyleId;
                        if (pStyle != null && pStyle.Val != null && pStyle.Val.HasValue)
                        {
                            writer.Write(" class='");
                            writer.Write(pStyle.Val.Value);
                            writer.Write('\'');
                        }
                        if (oxeProperties is Bold)
                        {
                            writer.Write(" style='");
                            writer.Write("font-weight:bold");
                            writer.Write("'");
                        }
                    }
                writer.Write(">");

                foreach (OpenXmlElement oxe2 in p)
                {
                    Run run = oxe2 as Run;

                    if (run == null)
                        continue;
                    writer.Write("<span ");
                    if (run.RunProperties != null)
                        foreach (var oxeProperties in run.RunProperties)
                            if (oxeProperties is Bold)
                            {
                                writer.Write(" style='");
                                writer.Write("font-weight:bold");
                                writer.Write("'");
                            }
                    writer.Write(">");
                    foreach (OpenXmlElement oxe3 in run)
                    {
                        if (oxe3 is Text)
                        {
                            SearchExtractor.Extract(oxe3.InnerText, ((FileStream)((StreamWriter)writer).BaseStream).Name, currentHeading);
                            writer.Write(oxe3.InnerText);
                        }
                        if (oxe3 is Picture)
                        {
                            Picture pict = (Picture)oxe3;
                            var shape = pict.GetFirstChild<Shape>();
                            if (shape != null)
                            {
                                writer.Write("<img ");
                                if (shape.Style != null && shape.Style.HasValue)
                                {
                                    writer.Write("style='");
                                    writer.Write(shape.Style.Value);
                                    writer.Write("'");
                                }
                                writer.Write(" src='");
                                writer.Write(doc.MainDocumentPart.GetPartById(((ImageData)shape.FirstChild).RelationshipId).Uri.ToString().Replace("/word/media/", "Images/") + "' />");
                            }
                        }
                        if (oxe3 is Drawing)
                        {
                            Drawing drawing = (Drawing)oxe3;
                            string id;
                            id = drawing.FirstChild
                                        .OfType<DocumentFormat.OpenXml.Drawing.Graphic>()
                                        .SelectMany(g => g.GraphicData
                                            .OfType<DocumentFormat.OpenXml.Drawing.Pictures.Picture>())
                                        .Select(pic => pic.BlipFill.Blip.Embed.Value)
                                        .FirstOrDefault();
                            writer.Write("<img ");
                            if (drawing.Anchor != null)
                            {
                                writer.Write("style='");
                                bool floating = false;
                                if ((int.Parse(drawing.Anchor.HorizontalPosition.PositionOffset.Text) + drawing.Anchor.Extent.Cx * 96 / 914400) > 8064)
                                {
                                    StyleExtractor.Write(writer, "float", "right");
                                    floating = true;
                                }
                                if (!floating)
                                    Write(writer, drawing.Anchor.HorizontalPosition);

                                Write(writer, drawing.Anchor.VerticalPosition);

                                DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapSquare wrap = drawing.Anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapSquare>();
                                if (wrap != null && wrap.WrapText != null)
                                    switch (wrap.WrapText.Value)
                                    {
                                        case DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapTextValues.Left:
                                            StyleExtractor.Write(writer, "float", "left");
                                            break;
                                        case DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapTextValues.Right:
                                            StyleExtractor.Write(writer, "float", "right");
                                            break;
                                    }
                                writer.Write("' ");

                            }
                            writer.Write(" src='");
                            writer.Write(doc.MainDocumentPart.GetPartById(id).Uri.ToString().Replace("/word/media/", "Images/") + "' />");
                        }
                        if (oxe3 is Break)
                            writer.Write("<br />");
                    }
                    writer.Write("</span>");
                }
                if (p.ParagraphProperties != null && p.ParagraphProperties.NumberingProperties != null)
                    writer.Write("</li>");

                if (p.ParagraphProperties != null && p.ParagraphProperties.NumberingProperties == null)
                    writer.WriteLine("</p>");
                //else
                //    writer.Write("</ul>");
            }

            writer.Flush();
            writer.Close();
        }

        private static void Write(TextWriter writer, DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalPosition horizontalPosition)
        {
            switch (horizontalPosition.RelativeFrom.Value)
            {
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.Character:
                    break;
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.Column:
                    StyleExtractor.Write(writer, "margin-left", int.Parse(horizontalPosition.PositionOffset.Text) * 96 / 914400 + "px");
                    break;
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.InsideMargin:
                    break;
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.LeftMargin:
                    break;
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.Margin:
                    break;
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.OutsideMargin:
                    break;
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.Page:
                    break;
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.RightMargin:
                    break;
                default:
                    break;
            }
        }

        private static void Write(TextWriter writer, DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalPosition verticalPosition)
        {
            switch (verticalPosition.RelativeFrom.Value)
            {
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.BottomMargin:
                    break;
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.InsideMargin:
                    break;
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.Line:
                    break;
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.Margin:
                    break;
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.OutsideMargin:
                    break;
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.Page:
                    break;
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.Paragraph:
                    StyleExtractor.Write(writer, "margin-top", (int.Parse(verticalPosition.PositionOffset.Text) * 96 / 914400) + "px");
                    break;
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.TopMargin:
                    break;
                default:
                    break;
            }
        }

        private static TextWriter CreateTextFile(string pageFile, bool isHtml)
        {
            return CreateTextFile(ref pageFile, isHtml);
        }

        private static TextWriter CreateTextFile(ref string pageFile, bool isHtml)
        {
            pageFile = GetPageFile(pageFile);
            var page = new StreamWriter(File.OpenWrite(pageFile + (isHtml ? ".html" : ".js")), Encoding.UTF8);
            if (isHtml)
                page.Write("<!DOCTYPE html><html><head><link href='CSS/generated.css' rel='stylesheet' type='text/css' /><link href='jquery-ui.css' rel='stylesheet' type='text/css' /><link href='../style.css' rel='stylesheet' type='text/css' /><script type='text/javascript' src='../jquery.js'><script type='text/javascript' src='jquery-ui.js'></script><script type='text/javascript' src='../jstree.js'></script></head><body>");
            return page;
        }

        private static string GetPageFile(string pageFile)
        {
            pageFile = pageFile.Replace('/', '-').Replace('\\', '-').Replace(":", "").Replace("« ", "").Replace(" »", "").Replace("?", "").Replace("!", "").Replace("’", "'").Replace("'", "").Trim();
            pageFile = pageFile.Substring(0, Math.Min(50, pageFile.Length));
            return pageFile;
        }

        internal static void ExtractToc()
        {
            XmlDocument doc = new XmlDocument();

            using (var tocXml = File.OpenRead("toc.xml"))
            {
                doc.Load(tocXml);
            }
            using (var toc = CreateTextFile("toc", false))
            {
                toc.WriteLine("$(function(){$('div#toc').tree({source:");
                RenderTocNodes(toc, ((XmlNode)doc.DocumentElement).ChildNodes);
                toc.WriteLine("});});");
            }
        }

        private static void RenderTocNodes(TextWriter toc, XmlNodeList xmlNodeList)
        {
            toc.Write("[");
            bool isFirst = true;
            foreach (XmlNode level in xmlNodeList)
            {
                if (isFirst)
                    isFirst = false;
                else
                    toc.Write(',');
                string text = level.Attributes["text"].Value;
                text = text.Substring(text.IndexOf('.') + 1);
                toc.Write("{label:'");
                toc.Write("<a href=\\'Doc/" + GetPageFile(text) + ".html\\' target=\\'content\\'>");
                toc.Write(text.Replace("'", "\\'"));
                toc.Write("</a>");
                toc.Write("'");
                if (level.HasChildNodes)
                {
                    toc.Write(", children:");
                    RenderTocNodes(toc, level.ChildNodes);
                }
                toc.WriteLine("}");
            }
            toc.Write("]");
        }
    }
}
