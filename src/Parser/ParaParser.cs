using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.ExtendedProperties;
using Dox2Word.Logging;
using Dox2Word.Model;
using Dox2Word.Parser.Models;

namespace Dox2Word.Parser
{
    internal static class ParaParser
    {
        private static readonly Logger logger = Logger.Instance;

        public static List<IParagraph> Parse(DocPara? para)
        {
            var paragraphs = new List<IParagraph>();

            if (para == null)
                return paragraphs;

            Parse(paragraphs, para, TextRunFormat.None);

            paragraphs.LastOrDefault()?.TrimTrailingWhitespace();

            return paragraphs;
        }

        private static void Parse(List<IParagraph> paragraphs, DocPara? para, TextRunFormat format, TextParagraphAlignment alignment = TextParagraphAlignment.Default)
        {
            if (para == null)
                return;

            foreach (object? part in para.Parts)
            {
                switch (part)
                {
                    case string s:
                        AddTextRun(paragraphs, alignment, s, format);
                        break;
                    case DocSimpleSect s when s.Kind == DoxSimpleSectKind.Warning:
                        Add(paragraphs, new TextParagraph(TextParagraphType.Warning));
                        Parse(paragraphs, s.Para, format);
                        Add(paragraphs, new TextParagraph());
                        break;
                    case OrderedList o:
                        ParseList(paragraphs, o, ListParagraphType.Number, format);
                        break;
                    case UnorderedList u:
                        ParseList(paragraphs, u, ListParagraphType.Bullet, format);
                        break;
                    case DocXRefSect r:
                        if (r.Title.FirstOrDefault() != "Todo")
                        {
                            foreach (var docPara in r.Description.Para)
                            {
                                Parse(paragraphs, docPara, format);
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
                    case BoldMarkup b:
                        Parse(paragraphs, b, format | TextRunFormat.Bold);
                        break;
                    case ItalicMarkup i:
                        Parse(paragraphs, i, format | TextRunFormat.Italic);
                        break;
                    case MonospaceMarkup m:
                        Parse(paragraphs, m, format | TextRunFormat.Monospace);
                        break;
                    case Ref r:
                        AddTextRun(paragraphs, alignment, r.Name, format | TextRunFormat.Monospace, r.RefId);
                        break;
                    case DocTable t:
                        ParseTable(paragraphs, t);
                        break;
                    case XmlElement e:
                        AddTextRun(paragraphs, alignment, e.InnerText, format);
                        break;
                    case DocSimpleSect:
                        break; // Ignore
                    default:
                        logger.Warning($"Unexpected text {part} ({part.GetType()}). Ignoring");
                        break;
                }
            }
        }

        private static T Add<T>(List<IParagraph> paragraphs, T paragraph) where T : IParagraph
        {
            // Trim the previous paragraph
            paragraphs.LastOrDefault()?.TrimTrailingWhitespace();

            paragraphs.Add(paragraph);
            return paragraph;
        }

        private static void Add(List<IParagraph> paragraphs, TextParagraphAlignment alignment, TextRun textRun)
        {
            if (paragraphs.LastOrDefault() is not TextParagraph paragraph)
            {
                paragraph = Add(paragraphs, new TextParagraph(alignment: alignment));
            }
            paragraph.Add(textRun);
        }

        private static void AddTextRun(List<IParagraph> paragraphs, TextParagraphAlignment alignment, string text, TextRunFormat format, string? referenceId = null) =>
            Add(paragraphs, alignment, new TextRun(text.TrimStart('\n'), format, referenceId));

        private static void ParseList(List<IParagraph> paragraphs, DocList docList, ListParagraphType type, TextRunFormat format)
        {
            var list = Add(paragraphs, new ListParagraph(type));
            foreach (var item in docList.Items)
            {
                // It could be that the para contains a warning or something which will add
                // another paragraph. In this case, we'll just ignore it.
                var paragraphList = new List<IParagraph>();
                foreach (var para in item.Paras)
                {
                    Parse(paragraphList, para, format);
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

        private static void ParseTable(List<IParagraph> paragraphs, DocTable table)
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
                        Parse(cellDoc.Paragraphs, para, TextRunFormat.None, alignment);
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
