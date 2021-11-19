using System;
using System.IO;
using System.Reflection;
using Dox2Word.Generator;
using Dox2Word.Logging;
using Dox2Word.Parser;

namespace Dox2Word
{
    public class Program
    {
        public static int Main(string[] args)
        {
            if (args.Length != 3)
            {
                var version = typeof(Program).Assembly.GetName().Version;
                Console.WriteLine($"Dox2Word version v{version.ToString(1)} (https://github.com/canton7/Dox2Word)");
                Console.WriteLine();

                string path = Path.GetFileNameWithoutExtension(Assembly.GetEntryAssembly().Location);
                Console.WriteLine($"Usage: {path} xml-folder Template.docx Output.docx");
                Console.WriteLine();
                Console.WriteLine($"xml-folder     Path to the folder where Doxygen wrote its XML output");
                Console.WriteLine($"Template.docx  Template word document. Must contain a placeholder with the text '{WordGenerator.Placeholder}'");
                Console.WriteLine($"Output.docx    Path to write the generated output to");
                return 254;
            }

            string xmlFolder = args[0];
            string templatePath = args[1];
            string outputPath = args[2];

            try
            {
                var project = XmlParser.Parse(xmlFolder);

                File.Delete(outputPath);
                File.Copy(templatePath, outputPath);
                using var output = File.Open(outputPath, FileMode.Open);
                WordGenerator.Generate(output, project);
            }
            catch (Exception e)
            {
                Logger.Instance.Error(e);
            }

            var logger = Logger.Instance;
            if (logger.HasErrors)
                return 1;
            if (logger.HasWarnings)
                return 2;
            return 0;
        }
    }
}
