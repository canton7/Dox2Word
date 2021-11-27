using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Xml;
using Dox2Word.Logging;
using Dox2Word.Model;
using Dox2Word.Parser.Models;

namespace Dox2Word.Parser
{
    internal class MixedModeParser
    {
        private static readonly Logger logger = Logger.Instance;
        private readonly Index index;

        public MixedModeParser(Index index)
        {
            this.index = index;
        }

        public List<IParagraph> Parse(List<object>? parts)
        {
            var paragraphs = new List<IParagraph>();

            if (parts == null)
                return paragraphs;

            this.Parse(paragraphs, parts, default);

            TrimLast(paragraphs);

            return paragraphs;
        }

        private void Parse(List<IParagraph> paragraphs, List<object>? parts, TextRunProperties properties, TextParagraphAlignment alignment = TextParagraphAlignment.Default)
        {
            if (parts == null)
                return;

            foreach (object? part in parts)
            {
                switch (part)
                {
                    case string s:
                        AddTextRun(paragraphs, alignment, s, properties);
                        break;
                    case DocEmoji e:
                        // Doxgen doesn't correctly keep ZWJ's, see https://github.com/doxygen/doxygen/issues/8918
                        AddTextRun(paragraphs, alignment, WebUtility.HtmlDecode(e.Unicode), properties); 
                        break;
                    case DocSimpleSect s when s.Kind == DoxSimpleSectKind.Warning:
                        Add(paragraphs, new TextParagraph(TextParagraphType.Warning));
                        this.Parse(paragraphs, s.Para?.Parts, properties);
                        Add(paragraphs, new TextParagraph());
                        break;
                    case DocList l:
                        this.ParseList(paragraphs, l, l.Type, properties);
                        break;
                    case DocXRefSect r:
                        if (r.Title.FirstOrDefault() != "Todo")
                        {
                            foreach (var docPara in r.Description.Para)
                            {
                                this.Parse(paragraphs, docPara.Parts, properties);
                            }
                        }
                        break;
                    case Listing l:
                        ParseListing(paragraphs, l);
                        break;
                    case Dot d:
                        Add(paragraphs, new DotParagraph(d.Contents));
                        break;
                    case DotFile d:
                        Add(paragraphs, new DotParagraph(File.ReadAllText(d.Name!)) { Caption = d.Contents });
                        break;
                    case DocUrlLink l:
                        this.ParseUrlLink(paragraphs, alignment, l, properties);
                        break;
                    case DocMarkup m:
                        if (m.Center)
                        {
                            Add(paragraphs, new TextParagraph(TextParagraphAlignment.Center));
                        }
                        this.Parse(paragraphs, m.Parts, properties.Combine(m.Properties));
                        if (m.Center)
                        {
                            Add(paragraphs, new TextParagraph());
                        }
                        break;
                    case Ref r:
                        AddTextRun(paragraphs, alignment, r.Name, properties.Combine(TextRunFormat.Monospace), this.index.ShouldReference(r.RefId) ? r.RefId : null);
                        break;
                    case DocTable t:
                        this.ParseTable(paragraphs, t);
                        break;
                    case XmlElement e:
                        logger.Warning($"Unrecognised text part {e.Name}. Taking raw string content");
                        AddTextRun(paragraphs, alignment, e.InnerText, properties);
                        break;
                    case DocSimpleSect:
                        break; // Ignore
                    default:
                        logger.Warning($"Unexpected text {part} ({part.GetType()}). Ignoring");
                        break;
                }
            }
        }

        private static T Add<T>(ICollection<IParagraph> paragraphs, T paragraph) where T : IParagraph
        {
            TrimLast(paragraphs);

            paragraphs.Add(paragraph);
            return paragraph;
        }

        private static void TrimLast(ICollection<IParagraph> paragraphs)
        {
            // Trim the previous paragraph
            if (paragraphs.LastOrDefault() is { } last)
            {
                last.TrimTrailingWhitespace();
                if (last.IsEmpty)
                {
                    paragraphs.Remove(last);
                }
            }
        }

        private static void Add(ICollection<IParagraph> paragraphs, TextParagraphAlignment alignment, IParagraphElement element)
        {
            if (paragraphs.LastOrDefault() is not TextParagraph paragraph)
            {
                paragraph = Add(paragraphs, new TextParagraph(alignment: alignment));
            }
            // If we're inserting text straight after a stand-alone line break (from newline), trim the leading space -- Doxygen seems to insert it for some reason
            if (element is TextRun textRun && paragraph.LastOrDefault() is TextRun previousTextRun && previousTextRun.Text == "\n")
            {
                textRun.Text = textRun.Text.TrimStart(' ');
            }
            paragraph.Add(element);
        }

        private static void AddTextRun(List<IParagraph> paragraphs, TextParagraphAlignment alignment, string text, TextRunProperties properties, string? referenceId = null) =>
            Add(paragraphs, alignment, new TextRun(text, properties, referenceId));

        private void ParseUrlLink(List<IParagraph> paragraphs, TextParagraphAlignment alignment, DocUrlLink urlLink, TextRunProperties properties)
        {
            var link = new UrlLink(urlLink.Url);
            var children = new List<IParagraph>();
            this.Parse(children, urlLink.Parts, properties);
            // We should only get one child, and it should be a TextParagraph
            if (children.Count > 0 && children[0] is TextParagraph textParagraph)
            {
                if (children.Count > 1)
                {
                    logger.Warning($"Unexpected number of children for DocUrlLink with URL {link.Url}");
                }

                link.AddRange(textParagraph.OfType<TextRun>());
                Add(paragraphs, alignment, link);
            }
            else
            {
                logger.Warning($"Unexpected content for DocUrlLink with URL {link.Url}");
            }
        }

        private void ParseList(List<IParagraph> paragraphs, DocList docList, ListParagraphType type, TextRunProperties properties)
        {
            var list = Add(paragraphs, new ListParagraph(type));
            foreach (var item in docList.Items)
            {
                // It could be that the para contains a warning or something which will add
                // another paragraph. In this case, we'll just ignore it.
                var paragraphList = new List<IParagraph>();
                foreach (var para in item.Paras)
                {
                    this.Parse(paragraphList, para.Parts, properties);
                }
                paragraphList.LastOrDefault()?.TrimTrailingWhitespace();
                list.Items.AddRange(paragraphList);
            }
        }

        private static void ParseListing(List<IParagraph> paragraphs, Listing listing)
        {
            var codeParagraph = Add(paragraphs, new CodeParagraph());
            foreach (var codeline in listing.Codelines)
            {
                var sb = new StringBuilder();
                foreach (var highlight in codeline.Highlights)
                {
                    foreach (object part in highlight.Parts)
                    {
                        switch (part)
                        {
                            case string s:
                                sb.Append(s);
                                break;
                            case Sp:
                                sb.Append(" ");
                                break;
                            case XmlElement e:
                                sb.Append(e.InnerText);
                                break;
                            default:
                                logger.Warning($"Unexpected code part {part} ({part.GetType()}). Ignoring");
                                break;
                        }
                    }
                }
                codeParagraph.Lines.Add(sb.ToString().TrimEnd(' '));
            }
        }

        private void ParseTable(List<IParagraph> paragraphs, DocTable table)
        {
            var tableDoc = new TableDoc()
            {
                NumColumns = table.NumCols,
            };
            foreach (var row in table.Rows)
            {
                var rowDoc = new TableRowDoc();
                tableDoc.Rows.Add(rowDoc);

                foreach (var entry in row.Entries)
                {
                    var cellDoc = new TableCellDoc()
                    {
                        ColSpan = entry.ColSpan,
                        RowSpan = entry.RowSpan,
                    };
                    rowDoc.Cells.Add(cellDoc);

                    var alignment = entry.Align switch
                    {
                        DoxAlign.Left => TextParagraphAlignment.Left,
                        DoxAlign.Center => TextParagraphAlignment.Center,
                        DoxAlign.Right => TextParagraphAlignment.Right,
                    };

                    foreach (var para in entry.Paras)
                    {
                        this.Parse(cellDoc.Paragraphs, para.Parts, default, alignment);
                        cellDoc.Paragraphs.LastOrDefault()?.TrimTrailingWhitespace();
                    }
                }
            }

            // Doxygen follows HTML's model where any cell can be a th. Word has a model where only the first row and
            // column can be headers. Simplify to Word's model. Check if all cells in the first row/column are headers
            tableDoc.FirstRowHeader = table.Rows.FirstOrDefault()?.Entries.All(x => x.IsTableHead == DoxBool.Yes) ?? false;
            tableDoc.FirstColumnHeader = table.Rows.All(x => x.Entries.FirstOrDefault()?.IsTableHead != DoxBool.No);

            Add(paragraphs, tableDoc);
        }
    }
}
