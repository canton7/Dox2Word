using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dox2Word.Generator
{
    public static class Extensions
    {
        public static Run FormatCode(this Run run)
        {
            run.RunProperties ??= new RunProperties();
            run.RunProperties.FontSize = new FontSize() { Val = "20" };
            run.PrependChild(new RunFonts() { Ascii = "Consolas" });
            return run;
        }

        public static Run FormatMiniHeading(this Run run)
        {
            run.RunProperties ??= new RunProperties();
            run.RunProperties.Bold = new Bold();
            return run;
        }

        public static Table AddBorders(this Table table)
        {
            var tableProperties = table.Elements<TableProperties>().FirstOrDefault() ??
                table.AppendChild(new TableProperties());
            tableProperties.TableBorders = new TableBorders(
                new TopBorder() { Val = BorderValues.Single },
                new RightBorder() { Val = BorderValues.Single },
                new BottomBorder() { Val = BorderValues.Single },
                new LeftBorder() { Val = BorderValues.Single },
                new InsideHorizontalBorder() { Val = BorderValues.Single },
                new InsideVerticalBorder() { Val = BorderValues.Single });

            return table;
        }

        public static void AppendRow(this Table table, string name, IEnumerable<Paragraph> value)
        {
            var row = table.AppendChild(new TableRow());

            var nameCell = row.AppendChild(new TableCell());
            var nameParagraph = nameCell.AppendChild(new Paragraph(new Run(new Text(name)).FormatCode()));
            nameParagraph.ParagraphProperties ??= new ParagraphProperties();
            nameParagraph.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines()
            {
                After = "0",
            };

            var valueCell = row.AppendChild(new TableCell());
            var valueList = value.ToList();
            valueCell.Append(valueList);

            if (valueList.LastOrDefault() is { } lastParagraph)
            {
                lastParagraph.ParagraphProperties ??= new ParagraphProperties();
                lastParagraph.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines()
                { 
                    After = "0",
                };
            }
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
