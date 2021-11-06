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

        public static Paragraph FormatWarning(this Paragraph paragraph)
        {
            paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph.ParagraphProperties.ParagraphBorders ??= new ParagraphBorders(new LeftBorder()
            {
                Val = BorderValues.Single,
                Space = 4,
                Size = 18,
                Color = "FFC000",
            });

            return paragraph;
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
                new InsideHorizontalBorder() { Val = BorderValues.Dotted },
                new InsideVerticalBorder() { Val = BorderValues.Dotted });

            return table;
        }

        public static void AppendRow(this Table table, string name, IEnumerable<Paragraph> value, string? inOut = null)
        {
            var row = table.AppendChild(new TableRow());

            if (inOut != null)
            {
                var inOutCell = row.AppendChild(new TableCell());
                var inOutParagraph = inOutCell.AppendChild(new Paragraph());
                var inOutRun = inOutParagraph.AppendChild(new Run(new Text(inOut)).FormatCode());
                inOutRun.RunProperties!.Italic = new Italic();
                Format(inOutParagraph, after: true);
            }

            var nameCell = row.AppendChild(new TableCell());
            var nameParagraph = nameCell.AppendChild(new Paragraph(new Run(new Text(name)).FormatCode()));
            Format(nameParagraph, after: true);

            var valueCell = row.AppendChild(new TableCell());
            var valueList = value.ToList();
            valueCell.Append(valueList);

            for (int i = 0; i < valueList.Count; i++)
            {
                Format(valueList[i], after: i == valueList.Count - 1);
            }

            static void Format(Paragraph paragraph, bool after)
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
