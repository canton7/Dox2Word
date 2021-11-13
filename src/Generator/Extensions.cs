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
            paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = styleId };

            return paragraph;
        }

        public static Run ApplyStyle(this Run run, string styleId)
        {
            run.RunProperties ??= new RunProperties();
            run.RunProperties.RunStyle = new RunStyle() { Val = styleId };

            return run;
        }

        public static Table ApplyStyle(this Table table, string styleId)
        {
            var tableProperties = table.Elements<TableProperties>().FirstOrDefault() ??
                table.AppendChild(new TableProperties());
            tableProperties.TableStyle = new TableStyle() { Val = styleId };

            return table;
        }

        public static Run WithProperties(this Run run, Action<RunProperties> updater)
        {
            run.RunProperties ??= new RunProperties();
            updater(run.RunProperties);

            return run;
        }

        public static Paragraph LeftAlign(this Paragraph paragraph)
        {
            paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph.ParagraphProperties.Justification = new Justification() { Val = JustificationValues.Left };
            return paragraph;
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
                paragraph.ParagraphProperties ??= new ParagraphProperties();
                paragraph.ParagraphProperties.KeepNext = new KeepNext();
                if (after)
                {
                    paragraph.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines()
                    {
                        After = "0",
                    };
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
