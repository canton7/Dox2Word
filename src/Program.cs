using System;
using Dox2Word.Parser;

namespace Dox2Word
{
    public class Program
    {
        public static void Main(string[] args)
        {
            new XmlParser("../../../test/C/Docs/xml").Parse();
        }
    }
}
