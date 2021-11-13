using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dox2Word.Generator
{
    public static class Extensions
    {
        public static Paragraph ApplyStyle(this Paragraph paragraph, string styleId)
        {
            return paragraph.WithProperties(x => x.ParagraphStyleId = new ParagraphStyleId() { Val = styleId });
        }

        public static Run ApplyStyle(this Run run, string styleId)
        {
            return run.WithProperties(x => x.RunStyle = new RunStyle() { Val = styleId });
        }

        public static Table ApplyStyle(this Table table, string styleId)
        {
            return table.WithProperties(x => x.TableStyle = new TableStyle() { Val = styleId });
        }

        public static Paragraph WithProperties(this Paragraph paragraph, Action<ParagraphProperties> updater)
        {
            if (paragraph.ParagraphProperties == null)
            {
                var p = new ParagraphProperties();
                updater(p);
                if (p.HasChildren)
                {
                    paragraph.ParagraphProperties = p;
                }
            }
            else
            {
                updater(paragraph.ParagraphProperties);
            }
            return paragraph;
        }

        public static Run WithProperties(this Run run, Action<RunProperties> updater)
        {
            if (run.RunProperties == null)
            {
                var p = new RunProperties();
                updater(p);
                if (p.HasChildren)
                {
                    run.RunProperties = p;
                }
            }
            else
            {
                updater(run.RunProperties);
            }
            return run;
        }

        public static Table WithProperties(this Table table, Action<TableProperties> updater)
        {
            var tableProperties = table.Elements<TableProperties>().FirstOrDefault();
            if (tableProperties == null)
            {
                tableProperties = new TableProperties();
                updater(tableProperties);
                if (tableProperties.HasChildren)
                {
                    table.AppendChild(tableProperties);
                }
            }
            else
            {
                updater(tableProperties);
            }
            return table;
        }

        public static TableCell WithProperties(this TableCell tableCell, Action<TableCellProperties> updater)
        {
            if (tableCell.TableCellProperties == null)
            {
                var p = new TableCellProperties();
                updater(p);
                if (p.HasChildren)
                {
                    tableCell.TableCellProperties = p;
                }
            }
            else
            {
                updater(tableCell.TableCellProperties);
            }
            return tableCell;
        }

        public static Paragraph LeftAlign(this Paragraph paragraph)
        {
            return paragraph.WithProperties(x => x.Justification = new Justification() { Val = JustificationValues.Left });
        }

        public static Table AddColumns(this Table table, int numColumns)
        {
            var tableGrid = table.AppendChild(new TableGrid());
            for (int i = 0; i < numColumns; i++)
            {
                tableGrid.AppendChild(new GridColumn());
            }
            return table;
        }

        public static void AppendRow(this Table table, string name, IEnumerable<OpenXmlElement> value, string? inOut = null)
        {
            var row = table.AppendChild(new TableRow());

            if (inOut != null)
            {
                var inOutCell = row.AppendChild(new TableCell());
                var inOutParagraph = inOutCell.AppendChild(new Paragraph());
                var inOutRun = inOutParagraph.AppendChild(new Run(new Text(inOut)).ApplyStyle(StyleManager.CodeCharStyleId));
                inOutRun.WithProperties(x => x.Italic = new Italic());
                inOutParagraph.FormatTableCellElement(after: true);
            }

            var nameCell = row.AppendChild(new TableCell());
            var nameParagraph = nameCell.AppendChild(new Paragraph(new Run(new Text(name)).ApplyStyle(StyleManager.CodeCharStyleId)));
            nameParagraph.FormatTableCellElement(after: true);

            var valueCell = row.AppendChild(new TableCell());
            var valueList = value.ToList();
            valueCell.Append(valueList);

            for (int i = 0; i < valueList.Count; i++)
            {
                valueList[i].FormatTableCellElement(after: i == valueList.Count - 1);
            }
        }

        public static OpenXmlElement FormatTableCellElement(this OpenXmlElement element, bool after)
        {
            if (element is Paragraph paragraph)
            {
                paragraph.WithProperties(x => x.KeepNext = new KeepNext());
                if (after)
                {
                    paragraph.WithProperties(x => x.SpacingBetweenLines = new SpacingBetweenLines() { After = "0" });
                }
            }

            return element;
        }

        public static TResult MaxOrDefault<TSource, TResult>(this IEnumerable<TSource> source, Func<TSource, TResult> selector, TResult defaultValue)
        {
            if (source is null)
                throw new ArgumentNullException(nameof(source));
            if (selector is null)
                throw new ArgumentNullException(nameof(selector));

            using var enumerator = source.GetEnumerator();
            if (!enumerator.MoveNext())
            {
                return defaultValue;
            }

            TResult currentMax = selector(enumerator.Current);
            while (enumerator.MoveNext())
            {
                var val = selector(enumerator.Current);
                if (Comparer<TResult>.Default.Compare(currentMax, val) < 0)
                {
                    currentMax = val;
                }
            }

            return currentMax;
        }
    }
}
