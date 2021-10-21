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
