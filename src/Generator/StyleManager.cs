using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dox2Word.Generator
{
    public class StyleManager
    {
        public const string TableStyleId = "DoxTable";
        public const string ListParagraphStyleId = "ListParagraph";

        public static void EnsureStyles(StyleDefinitionsPart stylesPart)
        {
            AddIfNotExists(stylesPart, TableStyleId, StyleValues.Table, "DoxTableStyle.xml");

            AddIfNotExists(stylesPart, ListParagraphStyleId, StyleValues.Paragraph, "List Paragraph", style =>
            {
                // In order to get list items next to each other, but still have space under the list, we
                // apply the same style to each, and set "Don't add space between paragraphs of the same style"
                // (ContextualSpacing). Since we apply the same style to each, this works. This is what Word does: if
                // the document already has a list, Word will have inserted this style; if not, we need to do it
                // ourselves.
                style.StyleParagraphProperties = new StyleParagraphProperties()
                {
                    ContextualSpacing = new ContextualSpacing(),
                };
            });
        }

        private static void AddIfNotExists(StyleDefinitionsPart stylesPart, string styleId, StyleValues styleType, string filename)
        {
            if (!stylesPart.Styles!.Elements<Style>().Any(x => x.StyleId == styleId))
            {
                var style = stylesPart.Styles.AppendChild(new Style()
                {
                    StyleId = styleId,
                    Type = styleType,
                });
                using var sr = new StreamReader(typeof(StyleManager).Assembly.GetManifestResourceStream($"Dox2Word.Generator.{filename}"));
                style.InnerXml = sr.ReadToEnd();
            }
        }

        private static void AddIfNotExists(StyleDefinitionsPart stylesPart, string styleId, StyleValues styleType, string styleName, Action<Style> configurer)
        {
            if (!stylesPart.Styles!.Elements<Style>().Any(x => x.StyleId == styleId))
            {
                var style = stylesPart.Styles.AppendChild(new Style()
                {
                    StyleId = styleId,
                    StyleName = new StyleName() { Val = styleName },
                    Type = styleType,
                });
                configurer(style);
            }
        }
    }
}
