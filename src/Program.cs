using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Dox2Word.Generator;
using Dox2Word.Logging;
using Dox2Word.Parser;
using Mono.Options;

namespace Dox2Word
{
    public class Options
    {
        public Dictionary<string, string> Placeholders { get; } = new();
    }

    public class Program
    {
        public static int Main(string[] args)
        {
            bool showHelp = false;
            string? xmlFolder = null;
            string? templatePath = null;
            string? outputPath = null;
            var options = new Options();

            var optionSet = new OptionSet();
            optionSet.Add("help|h", "Show this help", x => showHelp = x != null);
            optionSet.AddRequired("input=|i=", "'xml' folder created by Doxygen to use as input", x => xmlFolder = x);
            optionSet.AddRequired("template=|t=", $"Template word document. Must contain a placeholder paragraph with the text '<{WordGenerator.Placeholder}>'",
                x => templatePath = x);
            optionSet.AddRequired("output=|o=", "Path to write the generated document to", x => outputPath = x);
            optionSet.Add("placeholder==|p==", "{0:Placeholder} and {1:value} to substitute", (k, v) =>
            {
                if (k == null)
                    throw new OptionException("Missing name for placeholder", "-D");
                options.Placeholders[k] = v;
            });

            try
            {
                var remaining = optionSet.Parse(args);
                if (args.Length == 0 || showHelp)
                {
                    ShowHelp(optionSet);
                    return 0;
                }

                optionSet.ValidateRequiredOptions();
                if (remaining.Count > 0)
                    throw new OptionException($"Unrecognised option '{remaining[0]}'", remaining[0]);
            }
            catch (OptionException e)
            {
                Console.WriteLine($"{e.Message}. See --help for help");
                return 254;
            }

            try
            {
                var project = XmlParser.Parse(xmlFolder!);

                Logger.Instance.Info("Loading template");
                File.Delete(outputPath);
                File.Copy(templatePath, outputPath);
                using var output = File.Open(outputPath, FileMode.Open);
                WordGenerator.Generate(output, project, options);
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

        private static void ShowHelp(OptionSet optionSet)
        {
            var version = typeof(Program).Assembly.GetName().Version;
            Console.WriteLine($"Dox2Word version v{version.ToString(1)} (https://github.com/canton7/Dox2Word)");
            Console.WriteLine();

            string path = Path.GetFileNameWithoutExtension(Assembly.GetEntryAssembly().Location);
            Console.WriteLine($"Usage: {path} -i path/to/xml -t Template.docx -o Output.docx [-p placeholder=value[, ...]]");
            Console.WriteLine();
            Console.WriteLine("The available options are:");
            optionSet.WriteOptionDescriptions(Console.Out);
        }
    }
}
