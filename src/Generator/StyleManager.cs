using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dox2Word.Generator
{
    public class StyleManager
    {
        public const string TableStyleId = "DoxTable";
        public const string ParameterTableStyleId = "DoxParameterTable";
        public const string MiniHeadingStyleId = "DoxMiniHeading";
        public const string ParHeadingStyleId = "DoxParHeading";
        public const string CodeListingStyleId = "DoxCodeListing";
        public const string CodeStyleId = "DoxCode";
        public const string CodeCharStyleId = "DoxCodeChar";
        public const string WarningStyleId = "DoxWarning";
        public const string NoteStyleId = "DoxNote";
        public const string BlockQuoteStyleId = "DoxBlockQuote";
        public const string HyperlinkStyleId = "Hyperlink";
        public const string ListParagraphStyleId = "ListParagraph";

        private readonly StyleDefinitionsPart stylesPart;

        public int DefaultFontSize { get; }

        public StyleManager(StyleDefinitionsPart stylesPart)
        {
            this.stylesPart = stylesPart;

            this.DefaultFontSize = this.stylesPart.Styles!.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.FontSize is { } size
                ? int.Parse(size.Val)
                : 22;
        }

        public void EnsureStyles()
        {
            this.AddIfNotExists(TableStyleId, StyleValues.Table, "DoxTableStyle.xml");

            this.AddIfNotExists(ParameterTableStyleId, StyleValues.Table, "Dox Parameter Table", style =>
            {
                style.BasedOn = new BasedOn() { Val = "TableNormal" };
                style.StyleTableProperties = new StyleTableProperties()
                {
                    TableBorders = new TableBorders()
                    {
                        TopBorder = new TopBorder() { Val = BorderValues.Single, Color = "999999", },
                        LeftBorder = new LeftBorder() { Val = BorderValues.Single, Color = "999999", },
                        BottomBorder = new BottomBorder() { Val = BorderValues.Single, Color = "999999", },
                        RightBorder = new RightBorder() { Val = BorderValues.Single, Color = "999999", },
                        InsideHorizontalBorder = new InsideHorizontalBorder() { Val = BorderValues.Dotted, Color = "999999", },
                        InsideVerticalBorder = new InsideVerticalBorder() { Val = BorderValues.Dotted, Color = "999999", },
                    }
                };
            });

            this.AddIfNotExists(MiniHeadingStyleId, StyleValues.Paragraph, "Dox Mini Heading", style =>
            {
                style.BasedOn = new BasedOn() { Val = "Normal" };
                style.PrimaryStyle = new PrimaryStyle();
                style.StyleRunProperties = new StyleRunProperties()
                {
                    Bold = new Bold(),
                    FontSize = new FontSize() { Val = (this.DefaultFontSize + 2).ToString() },
                };
                style.StyleParagraphProperties = new StyleParagraphProperties()
                {
                     SpacingBetweenLines = new SpacingBetweenLines()
                     {
                         LineRule = LineSpacingRuleValues.Auto,
                         Before = "160",
                         Line = "240",
                     },
                     KeepNext = new KeepNext(),
                };
            });

            this.AddIfNotExists(ParHeadingStyleId, StyleValues.Paragraph, "Dox Par Heading", style =>
            {
                style.BasedOn = new BasedOn() { Val = "Normal" };
                style.PrimaryStyle = new PrimaryStyle();
                style.StyleRunProperties = new StyleRunProperties()
                {
                    Bold = new Bold(),
                    FontSize = new FontSize() { Val = (this.DefaultFontSize + 1).ToString() },
                };
                style.StyleParagraphProperties = new StyleParagraphProperties()
                {
                    SpacingBetweenLines = new SpacingBetweenLines()
                    {
                        LineRule = LineSpacingRuleValues.Auto,
                        Before = "160",
                        Line = "240",
                    },
                    KeepNext = new KeepNext(),
                };
            });

            this.AddIfNotExists(CodeCharStyleId, StyleValues.Character, "Dox Code Char", style =>
            {
                style.BasedOn = new BasedOn() { Val = "DefaultParagraphFont" };
                style.LinkedStyle = new LinkedStyle() { Val = CodeStyleId };
                style.StyleRunProperties = new StyleRunProperties()
                {
                    FontSize = new FontSize() { Val = (this.DefaultFontSize - 2).ToString() },
                    RunFonts = new RunFonts() { Ascii = "Consolas" },
                };
            });

            this.AddIfNotExists(CodeStyleId, StyleValues.Paragraph, "Dox Code", style =>
            {
                style.BasedOn = new BasedOn() { Val = "Normal" };
                style.LinkedStyle = new LinkedStyle() { Val = CodeCharStyleId };
                style.PrimaryStyle = new PrimaryStyle();
                style.StyleParagraphProperties = new StyleParagraphProperties()
                {
                    KeepNext = new KeepNext(),
                    KeepLines = new KeepLines(),
                    SpacingBetweenLines = new SpacingBetweenLines()
                    {
                        Line = "240",
                        LineRule = LineSpacingRuleValues.Auto,
                    },
                    ContextualSpacing = new ContextualSpacing(),
                };
                style.StyleRunProperties = new StyleRunProperties()
                {
                    FontSize = new FontSize() { Val = (this.DefaultFontSize - 2).ToString() },
                    RunFonts = new RunFonts() { Ascii = "Consolas" },
                };
            });

            this.AddIfNotExists(CodeListingStyleId, StyleValues.Paragraph, "Dox Code Listing", style =>
            {
                style.BasedOn = new BasedOn() { Val = CodeStyleId };
                style.LinkedStyle = new LinkedStyle() { Val = CodeCharStyleId };
                style.PrimaryStyle = new PrimaryStyle();
                style.StyleParagraphProperties = new StyleParagraphProperties()
                {
                    ParagraphBorders = new ParagraphBorders()
                    {
                        TopBorder = new TopBorder() { Val = BorderValues.Single },
                        LeftBorder = new LeftBorder() { Val = BorderValues.Single },
                        BottomBorder = new BottomBorder() { Val = BorderValues.Single },
                        RightBorder = new RightBorder() { Val = BorderValues.Single },
                    },
                };
            });

            this.AddIfNotExists(WarningStyleId, StyleValues.Paragraph, "Dox Warning", style =>
            {
                style.BasedOn = new BasedOn() { Val = "Normal" };
                style.PrimaryStyle = new PrimaryStyle();
                style.StyleParagraphProperties = new StyleParagraphProperties()
                {
                    ParagraphBorders = new ParagraphBorders(new LeftBorder()
                    {
                        Val = BorderValues.Single,
                        Space = 4,
                        Size = 18,
                        Color = "FFC000",
                    }),
                };
            });

            this.AddIfNotExists(NoteStyleId, StyleValues.Paragraph, "Dox Note", style =>
            {
                style.BasedOn = new BasedOn() { Val = "Normal" };
                style.PrimaryStyle = new PrimaryStyle();
                style.StyleParagraphProperties = new StyleParagraphProperties()
                {
                    ParagraphBorders = new ParagraphBorders(new LeftBorder()
                    {
                        Val = BorderValues.Single,
                        Space = 4,
                        Size = 18,
                        Color = "D0C000",
                    }),
                };
            });

            this.AddIfNotExists(BlockQuoteStyleId, StyleValues.Paragraph, "Dox Block Quote", style =>
            {
                style.BasedOn = new BasedOn() { Val = "Normal" };
                style.PrimaryStyle = new PrimaryStyle();
                style.StyleParagraphProperties = new StyleParagraphProperties()
                {
                    Indentation = new Indentation() { Left = "567" },
                };
            });

            this.AddIfNotExists(HyperlinkStyleId, StyleValues.Character, "Hyperlink", style =>
            {
                style.BasedOn = new BasedOn() { Val = "DefaultParagraphFont" };
                style.UnhideWhenUsed = new UnhideWhenUsed();
                style.StyleRunProperties = new StyleRunProperties()
                {
                    Color = new Color() { ThemeColor = ThemeColorValues.Hyperlink, Val = "0563C1" },
                    Underline = new Underline() {  Val = UnderlineValues.Single },
                };
            });

            this.AddIfNotExists(ListParagraphStyleId, StyleValues.Paragraph, "List Paragraph", style =>
            {
                // In order to get list items next to each other, but still have space under the list, we
                // apply the same style to each, and set "Don't add space between paragraphs of the same style"
                // (ContextualSpacing). Since we apply the same style to each, this works. This is what Word does: if
                // the document already has a list, Word will have inserted this style; if not, we need to do it
                // ourselves.
                style.BasedOn = new BasedOn() { Val = "Normal" };
                style.StyleParagraphProperties = new StyleParagraphProperties()
                {
                    ContextualSpacing = new ContextualSpacing(),
                };
            });
        }

        private void AddIfNotExists(string styleId, StyleValues styleType, string filename)
        {
            if (!this.stylesPart.Styles!.Elements<Style>().Any(x => x.StyleId == styleId))
            {
                var style = this.stylesPart.Styles.AppendChild(new Style()
                {
                    StyleId = styleId,
                    Type = styleType,
                });
                using var sr = new StreamReader(typeof(StyleManager).Assembly.GetManifestResourceStream($"Dox2Word.Generator.{filename}"));
                style.InnerXml = sr.ReadToEnd();
            }
        }

        private void AddIfNotExists(string styleId, StyleValues styleType, string styleName, Action<Style> configurer)
        {
            if (!this.stylesPart.Styles!.Elements<Style>().Any(x => x.StyleId == styleId))
            {
                var style = this.stylesPart.Styles.AppendChild(new Style()
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
