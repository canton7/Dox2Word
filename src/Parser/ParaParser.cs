using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

            return paragraphs;
        }

        private static void Parse(List<IParagraph> paragraphs, DocPara? para, TextRunFormat format)
        {
            if (para == null)
                return;

            foreach (object? part in para.Parts)
            {
                switch (part)
                {
                    case string s:
                        AddTextRun(paragraphs, s, format);
                        break;
                    case DocSimpleSect s when s.Kind == DoxSimpleSectKind.Warning:
                        NewParagraph(paragraphs, ParagraphType.Warning);
                        Parse(paragraphs, s.Para, format);
                        NewParagraph(paragraphs);
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
                        AddTextRun(paragraphs, r.Name, format | TextRunFormat.Monospace);
                        break;
                    case DocTable t:
                        ParseTable(paragraphs, t);
                        break;
                    case XmlElement e:
                        AddTextRun(paragraphs, e.InnerText, format);
                        break;
                    case DocSimpleSect:
                        break; // Ignore
                    default:
                        logger.Warning($"Unexpected text {part} ({part.GetType()}). Ignoring");
                        break;
                };
            }
        }

        private static void NewParagraph(List<IParagraph> elements, ParagraphType type = ParagraphType.Normal) =>
            elements.Add(new TextParagraph(type));

        private static void Add(List<IParagraph> paragraphs, TextRun textRun)
{
            if (paragraphs.LastOrDefault() is not TextParagraph paragraph)
            {
                paragraph = new TextParagraph();
                paragraphs.Add(paragraph);
            }
            paragraph.Add(textRun);
        }

        private static void AddTextRun(List<IParagraph> paragraphs, string text, TextRunFormat format) =>
            Add(paragraphs, new TextRun(text.TrimStart('\n'), format));

        private static void ParseList(List<IParagraph> paragraphs, DocList docList, ListParagraphType type, TextRunFormat format)
        {
            var list = new ListParagraph(type);
            paragraphs.Add(list);
            foreach (var item in docList.Items)
            {
                // It could be that the para contains a warning or something which will add
                // another paragraph. In this case, we'll just ignore it.
                var paragraphList = new List<IParagraph>();
                foreach (var para in item.Paras)
                {
                    Parse(paragraphList, para, format);
                }
                list.Items.AddRange(paragraphList);
            }
        }

        private static void ParseListing(List<IParagraph> paragraphs, Listing listing)
        {
            var codeParagraph = new CodeParagraph();
            paragraphs.Add(codeParagraph);
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
                codeParagraph.Lines.Add(sb.ToString());
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

                    foreach (var para in entry.Paras)
                    {
                        Parse(cellDoc.Paragraphs, para, TextRunFormat.None);
                    }
                }
            }

            // Doxygen follows HTML's model where any cell can be a th. Word has a model where only the first row and
            // column can be headers. Simplify to Word's model. Check if all cells in the first row/column are headers
            tableDoc.FirstRowHeader = table.Rows.FirstOrDefault()?.Entries.All(x => x.IsTableHead == DoxBool.Yes) ?? false;
            tableDoc.FirstColumnHeader = table.Rows.All(x => x.Entries.FirstOrDefault()?.IsTableHead != DoxBool.No);

            paragraphs.Add(tableDoc);
        }
    }
}
