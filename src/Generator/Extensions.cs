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
    }
}
