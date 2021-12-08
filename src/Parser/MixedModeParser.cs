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
        private readonly string basePath;
        private readonly Index index;

        public MixedModeParser(string basePath, Index index)
        {
            this.basePath = basePath;
            this.index = index;
        }

        public List<IParagraph> Parse(List<object>? parts)
        {
            var paragraphs = new List<IParagraph>();

            if (parts == null)
                return paragraphs;

            this.Parse(paragraphs, parts, default, TextParagraphAlignment.Default);

            TrimLast(paragraphs);

            return paragraphs;
        }

        private void Parse(List<IParagraph> paragraphs, List<object>? parts, TextRunProperties properties, TextParagraphAlignment alignment)
        {
            if (parts == null)
                return;

            foreach (object? part in parts)
            {
                switch (part)
                {
                    case string s:
                        AddTextRun(paragraphs, s, properties, alignment);
                        break;
                    case DocAnchor a:
                        Add(paragraphs, alignment, new TextRun("", anchorId: a.Id));
                        break;
                    case DocParBlock b:
                        foreach (var para in b.Paras)
                        {
                            this.Parse(paragraphs, para.Parts, properties, alignment);
                            Add(paragraphs, new TextParagraph());
                        }
                        break;
                    case DocEmoji e:
                        // Doxgen doesn't correctly keep ZWJ's, see https://github.com/doxygen/doxygen/issues/8918
                        AddTextRun(paragraphs, WebUtility.HtmlDecode(e.Unicode), properties, alignment); 
                        break;
                    case DocSimpleSect s when s.Kind == DoxSimpleSectKind.Warning:
                        Add(paragraphs, new TextParagraph(TextParagraphType.Warning));
                        this.Parse(paragraphs, s.Para?.Parts, properties, alignment);
                        Add(paragraphs, new TextParagraph());
                        break;
                    case DocSimpleSect s when s.Kind == DoxSimpleSectKind.Note:
                        Add(paragraphs, new TextParagraph(TextParagraphType.Note));
                        this.Parse(paragraphs, s.Para?.Parts, properties, alignment);
                        Add(paragraphs, new TextParagraph());
                        break;
                    case DocSimpleSect s when s.Kind == DoxSimpleSectKind.Par:
                        if (s.Title.Parts.Count > 0)
                        {
                            Add(paragraphs, new TitleParagraph());
                            this.Parse(paragraphs, s.Title.Parts, default, alignment);
                        }
                        Add(paragraphs, new TextParagraph());
                        this.Parse(paragraphs, s.Para?.Parts, properties, alignment);
                        Add(paragraphs, new TextParagraph());
                        break;
                    case DocSimpleSect s when s.Kind == DoxSimpleSectKind.Return:
                        break; // Ignore
                    case DocSimpleSect s:
                        Add(paragraphs, new TextParagraph());
                        this.Parse(paragraphs, s.Para?.Parts, properties, alignment);
                        break;
                    case DocList l:
                        this.ParseList(paragraphs, l, l.Type, properties, alignment);
                        break;
                    case DocBlockQuote b:
                        foreach (var para in b.Paras)
                        {
                            Add(paragraphs, new TextParagraph(TextParagraphType.BlockQuote));
                            this.Parse(paragraphs, para.Parts, properties, alignment);
                        }
                        Add(paragraphs, new TextParagraph());
                        break;
                    case DocXRefSect r:
                        if (r.Title.FirstOrDefault() is not ("Todo" or "Bug"))
                        {
                            foreach (var docPara in r.Description.Para)
                            {
                                this.Parse(paragraphs, docPara.Parts, properties, alignment);
                            }
                        }
                        break;
                    case Listing l:
                        ParseListing(paragraphs, l);
                        break;
                    case PreformattedDocMarkup m:
                        Add(paragraphs, new TextParagraph(TextParagraphType.Preformatted));
                        this.Parse(paragraphs, m.Parts, properties, alignment);
                        Add(paragraphs, new TextParagraph());
                        break;
                    case Image i when i.Type == "html":
                        if (i.Inline != DoxBool.Yes)
                        {
                            Add(paragraphs, new TextParagraph());
                        }
                        Add(paragraphs, alignment, new ImageElement(Path.Combine(this.basePath, i.Name), i.Caption));
                        if (i.Inline != DoxBool.Yes)
                        {
                            Add(paragraphs, new TextParagraph());
                        }
                        break;
                    case Image:
                        break;
                    case Dot d:
                        Add(paragraphs, new DotParagraph(d.Contents) { Caption = d.Caption });
                        break;
                    case DotFile d:
                        Add(paragraphs, new DotParagraph(File.ReadAllText(d.Name!)) { Caption = d.Contents });
                        break;
                    case DocUrlLink l:
                        this.ParseUrlLink(paragraphs, l, properties, alignment);
                        break;
                    case DocMarkup m:
                        if (m.Center)
                        {
                            Add(paragraphs, new TextParagraph(TextParagraphAlignment.Center));
                        }
                        this.Parse(paragraphs, m.Parts, properties.Combine(m.Properties), alignment);
                        if (m.Center)
                        {
                            Add(paragraphs, new TextParagraph());
                        }
                        break;
                    case Ref r:
                        AddTextRun(paragraphs, r.Name, properties.Combine(TextRunFormat.Monospace), alignment, this.index.ShouldReference(r.RefId) ? r.RefId : null);
                        break;
                    case DocTable t:
                        this.ParseTable(paragraphs, t);
                        break;
                    case DocVariableList l:
                        this.ParseVariableList(paragraphs, l);
                        break;
                    case HorizontalRulerDocEmptyType:
                        if (paragraphs.LastOrDefault() is TextParagraph p)
                        {
                            p.HasHorizontalRuler = true;
                        }
                        else
                        {
                            Add(paragraphs, new TextParagraph() { HasHorizontalRuler = true });
                            Add(paragraphs, new TextParagraph());
                        }
                        break;
                    case XmlElement e:
                        logger.Warning($"Unrecognised text part {e.Name}. Taking raw string content");
                        AddTextRun(paragraphs, e.InnerText, properties, alignment);
                        break;
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
            if (paragraphs.LastOrDefault() is not ICollection<IParagraphElement> paragraph)
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

        private static void AddTextRun(List<IParagraph> paragraphs, string text, TextRunProperties properties, TextParagraphAlignment alignment, string? referenceId = null) =>
            Add(paragraphs, alignment, new TextRun(text, properties, referenceId));

        private void ParseUrlLink(List<IParagraph> paragraphs, DocUrlLink urlLink, TextRunProperties properties, TextParagraphAlignment alignment)
        {
            var link = new UrlLink(urlLink.Url);
            var textParagraph = this.ParseSingleParagraph(urlLink, properties, alignment);
            if (textParagraph != null)
            { 
                link.AddRange(textParagraph.OfType<TextRun>());
                Add(paragraphs, alignment, link);
            }
        }

        private void ParseList(List<IParagraph> paragraphs, DocList docList, ListParagraphType type, TextRunProperties properties, TextParagraphAlignment alignment)
        {
            var list = Add(paragraphs, new ListParagraph(type));
            foreach (var item in docList.Items)
            {
                // It could be that the para contains a warning or something which will add
                // another paragraph. In this case, we'll just ignore it.
                var paragraphList = new List<IParagraph>();
                foreach (var para in item.Paras)
                {
                    this.Parse(paragraphList, para.Parts, properties, alignment);
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
                Id = table.Caption?.Id,
                NumColumns = table.NumCols,
                Caption = table.Caption == null ? null : this.Parse(table.Caption.Parts)?.FirstOrDefault() as TextParagraph,
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

        private void ParseVariableList(List<IParagraph> paragraphs, DocVariableList list)
        {
            var listParagraph = new DefinitionListParagraph();

            // We expect a sequence of DocVarListEntry, DocListItem pairs (although this isn't enforced by the parser)
            for (int i = 0; i < list.Parts.Count;)
            {
                object entryObj = list.Parts[i];
                i++;

                if (entryObj is not DocVarListEntry entry)
                {
                    logger.Warning("Malformed variablelist: varlistentry not at expected location");
                    continue;
                }

                if (i >= list.Parts.Count)
                {
                    logger.Warning("Malformed variablelist: missing final listitem");
                    break;
                }

                object itemObj = list.Parts[i];
                if (itemObj is not DocListItem item)
                {
                    logger.Warning("Malformed variablelist: listitem not at expected location");
                    // Go around the loop again, where we'll try parsing it as a DocVarListEntry
                    continue;
                }

                i++;

                var termParagraph = this.ParseSingleParagraph(entry.Term, default, TextParagraphAlignment.Default) ?? new TextParagraph();
                var definitionListEntry = new DefinitionListEntry()
                {
                    Term = termParagraph,
                };
                foreach (var para in item.Paras)
                {
                    this.Parse(definitionListEntry.Description, para.Parts, default, TextParagraphAlignment.Default);
                }
                listParagraph.Entries.Add(definitionListEntry);
            }

            Add(paragraphs, listParagraph);
        }

        private TextParagraph? ParseSingleParagraph(DocTitleCmdGroup title, TextRunProperties properties, TextParagraphAlignment alignment)
        {
            var children = new List<IParagraph>();
            this.Parse(children, title.Parts, properties, alignment);
            if (children.Count > 0 && children[0] is TextParagraph textParagraph)
            {
                if (children.Count > 1)
                {
                    logger.Warning("Unexpected number of children for DocTitleCmdGroup");
                }

                return textParagraph;
            }

            logger.Warning("Unexpected content for DocTitleCmdGroup");
            return null;
        }
    }
}
