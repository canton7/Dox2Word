using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dox2Word.Generator;

namespace Dox2Word.Logging
{
    public class Logger
    {
        public static Logger Instance { get; } = new Logger();

        private Logger() { }

        public void Info(string text)
        {
            WriteLevel(null, "INFO");
            Console.WriteLine(text);
        }

        public void Warning(string text)
        {
            WriteLevel(ConsoleColor.Yellow, "WARNING");
            Console.WriteLine(text);
        }

        public void Error(Exception e)
        {
            WriteLevel(ConsoleColor.Red, "ERROR");
            Console.WriteLine(e.Message);

            Console.WriteLine(e.StackTrace);
        }

        private static void WriteLevel(ConsoleColor? color, string text)
        {
            color ??= Console.ForegroundColor;
            Console.Write("[");
            var oldColor = Console.ForegroundColor;
            Console.ForegroundColor = color.Value;
            Console.Write(text);
            Console.ForegroundColor = oldColor;
            Console.Write("] ");
        }
    }
}
