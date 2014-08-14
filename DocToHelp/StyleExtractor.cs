using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace DocToHelp
{
    static class StyleExtractor
    {
        internal static void Extract(DocumentFormat.OpenXml.Packaging.ThemePart themes, DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart styleDefinitionsPart)
        {
            using (TextWriter css = new StreamWriter(File.OpenWrite("CSS\\generated.css")))
            {
                css.Write("body{ ");
                Extract(styleDefinitionsPart.Styles.DocDefaults.ParagraphPropertiesDefault.ParagraphPropertiesBaseStyle, css);
                Extract(styleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle, css);
                Write(css, "font-family", themes.Theme.FirstChild.ChildElements.OfType<DocumentFormat.OpenXml.Drawing.FontScheme>().SelectMany(f => f.ChildElements).OfType<DocumentFormat.OpenXml.Drawing.MajorFont>().SelectMany(f => f.ChildElements).OfType<DocumentFormat.OpenXml.Drawing.LatinFont>().FirstOrDefault().Typeface + ", Arial");
                css.WriteLine("}");
                foreach (OpenXmlElement oxe in styleDefinitionsPart.Styles.ChildElements)
                {
                    Style style = oxe as Style;
                    if (style == null)
                        continue;
                    css.Write("." + style.StyleId.Value);
                    if (style.LinkedStyle != null && style.LinkedStyle.Val != null && style.LinkedStyle.Val.HasValue)
                        css.Write(", ." + style.LinkedStyle.Val.Value);
                    css.Write("{ ");
                    Extract(themes, css, style);
                    css.WriteLine("}");
                }
            }
        }

        private static void Extract(ThemePart theme, TextWriter css, Style style)
        {
            if (style.StyleRunProperties == null)
                return;
            Extract(style.StyleRunProperties, style.StyleParagraphProperties, css, theme);
        }

        private static void Extract(StyleRunProperties runStyle, StyleParagraphProperties pStyle, TextWriter css, ThemePart theme)
        {
            if (runStyle.Bold != null && (runStyle.Bold.Val == null || runStyle.Bold.Val.HasValue && runStyle.Bold.Val.Value))
                Write(css, "font-weight", "bold");
            Write(css, runStyle.Border);

            Write(css, "color", runStyle.Color);
            if (runStyle.Shading != null)
            {
                if (runStyle.Shading.Fill != null && runStyle.Shading.Fill.HasValue)
                    Write(css, "background-color", "#" + runStyle.Shading.Fill);
            }
            if (pStyle != null && pStyle.Justification != null)
                Write(css, "text-align", pStyle.Justification.Val);
            Write(css, "font-style", "italic", runStyle.Italic);
            if (runStyle.RunFonts != null)
                Write(css, "font-face", runStyle.RunFonts.Ascii);
            if (pStyle != null && pStyle.ParagraphBorders != null)
                Write(css, pStyle.ParagraphBorders);
            if (runStyle.Caps != null)
                Write(css, "text-transform", "uppercase");
        }

        private static void Write(TextWriter css, string p, string p2, OnOffType propertyValue)
        {
            if (propertyValue != null && (propertyValue.Val == null || propertyValue.Val.HasValue && propertyValue.Val.Value))
                Write(css, p, p2);
        }

        public static void Write(TextWriter writer, string propertyName, string propertyValue)
        {
            writer.Write(propertyName);
            writer.Write(':');
            writer.Write(propertyValue);
            writer.Write("; ");
        }

        private static void Write(TextWriter writer, string propertyName, OpenXmlSimpleType propertyValue)
        {
            if (propertyValue == null || !propertyValue.HasValue)
                return;
            Write(writer, propertyName, propertyValue.InnerText);
        }

        private static void Write(TextWriter writer, Border border)
        {
            if (border == null)
                return;
            Write(writer, "", border);
            if (border.HasChildren)
                foreach (BorderType b in border.ChildElements)
                {
                    Write(writer, "", b);
                }
        }

        private static void Write(TextWriter writer, ParagraphBorders border)
        {
            if (border == null)
                return;
            Write(writer, "top", border.TopBorder);
            Write(writer, "left", border.LeftBorder);
            Write(writer, "right", border.RightBorder);
            Write(writer, "bottom", border.BottomBorder);

        }

        private static void Write(TextWriter writer, string p, BorderType b)
        {
            if (b == null)
                return;
            if (!string.IsNullOrEmpty(p) && !p.StartsWith("-"))
                p = "-" + p;
            if (b.Color.HasValue)
                Write(writer, "border" + p + "-color", "#" + b.Color.Value);
            if (b.Size.HasValue)
                Write(writer, "border" + p + "-width", (Math.Max(20, b.Size.Value) / 20.0) + "pt");
            if (b.Val.HasValue)
            {
                switch (b.Val.Value)
                {
                    case BorderValues.Dotted:
                        Write(writer, "border" + p + "-style", "dotted");
                        break;
                    case BorderValues.None:
                        Write(writer, "border" + p + "-style", "none");
                        break;
                    case BorderValues.Single:
                        Write(writer, "border" + p + "-style", "solid");
                        break;
                    default:
                        Write(writer, "border" + p + "-style", b.Val.Value.ToString().ToLower());
                        break;
                }
            }
        }


        private static void Write(TextWriter writer, string p, uint p_2)
        {
            Write(writer, p, p_2.ToString());
        }

        private static void Write(TextWriter writer, string propertyName, BorderValues border)
        {
            switch (border)
            {
                case BorderValues.BasicBlackDashes:
                    Write(writer, propertyName, "dash");
                    break;
                case BorderValues.BasicBlackDots:
                    Write(writer, propertyName, "dot");
                    break;
                case BorderValues.BasicThinLines:
                    Write(writer, propertyName, "solid");
                    break;
            }
        }

        private static void Write(TextWriter writer, string propertyName, Color color)
        {
            if (color == null || color.Val == null || !color.Val.HasValue)
                return;
            Write(writer, propertyName, "#" + color.Val.Value);
        }

        static void Extract(this RunPropertiesBaseStyle style, TextWriter css)
        {
            if (style.RunFonts != null)
                Write(css, "font-face", style.RunFonts.Ascii);
            if (style.Bold != null && style.Bold.Val.HasValue && style.Bold.Val.Value)
                css.WriteLine("font-weight: bold");
            if (style.Border != null)
            {
                if (style.Border.Color.HasValue)
                    css.WriteLine("border-color:" + style.Border.Color.Value);
                if (style.Border.Size.HasValue)
                    css.WriteLine("border-width:" + style.Border.Size.Value);
                if (style.Border.Val.HasValue)
                {
                    css.Write("border-style:");
                    switch (style.Border.Val.Value)
                    {
                        case BorderValues.BasicBlackDashes:
                            css.WriteLine("dash");
                            break;
                        case BorderValues.BasicBlackDots:
                            css.WriteLine("dot");
                            break;
                        case BorderValues.BasicBlackSquares:
                            break;
                        case BorderValues.BasicThinLines:
                            css.WriteLine("solid");
                            break;
                    }
                }
            }
            if (style.Color != null && style.Color.Val.HasValue)
            {
                css.WriteLine("color:#" + style.Color.Val.Value);
            }
        }

        static void Extract(this ParagraphPropertiesBaseStyle style, TextWriter css)
        {
            if (style == null)
                return;
            if (style.Justification != null)
                Write(css, "text-align", style.Justification.Val);
        }
    }
}
