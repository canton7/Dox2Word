using System;
using System.IO;
using Dox2Word.Generator;
using Dox2Word.Parser;

namespace Dox2Word
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var project = new XmlParser("../../../test/C/Docs/xml").Parse();

            File.Delete("../../../test/C/Docs/Output.docx");
            File.Copy("../../../test/C/Docs/Template.docx", "../../../test/C/Docs/Output.docx");
            using var output = File.Open("../../../test/C/Docs/Output.docx", FileMode.Open);
            WordGenerator.Generate(output, project);
        }
    }
}
