using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Dox2Word.Model;

namespace Dox2Word.Generator
{
    public class WordGenerator
    {
        private const string Placeholder = "<INSERT HERE>";

        private readonly Stream stream;

        private readonly List<Paragraph> paragraphs = new();

        public WordGenerator(Stream stream)
        {
            this.stream = stream;
        }

        public void Generate(Project project)
        {
            this.paragraphs.Clear();

            using var doc = WordprocessingDocument.Open(this.stream, isEditable: true);
            var body = doc.MainDocumentPart!.Document.Body!;

            foreach (var group in project.Groups)
            {
                this.WriteGroup(group, 2);
            }

            foreach (var paragraph in this.paragraphs)
            {
                body.AppendChild(paragraph);
            }
        }

        private void WriteGroup(Group group, int headingLevel)
        {
            this.WriteHeading(group.Name, headingLevel);
            
            this.WriteDescriptions(group.Descriptions);

            this.WriteGlobalVariables(group.GlobalVariables, headingLevel + 1);
            this.WriteFunctions(group.Functions, headingLevel + 1);

            foreach (var subGroup in group.SubGroups)
            {
                this.WriteGroup(subGroup, headingLevel + 1);
            }
        }

        private void WriteGlobalVariables(List<Variable> variables, int headingLevel)
        {
            if (variables.Count == 0)
                return;

            this.WriteHeading("Global Variables", headingLevel);

            foreach (var variable in variables)
            {
                this.WriteHeading(variable.Name, headingLevel + 1);

                var paragraph = this.CreateParagraph();
                var run = paragraph.AppendChild(new Run(new Text(variable.Definition)));
                run.FormatCode();

                this.WriteDescriptions(variable.Descriptions);
            }
        }

        private void WriteFunctions(List<Function> functions, int headingLevel)
        {
            if (functions.Count == 0)
                return;

            this.WriteHeading("Functions", headingLevel);

            foreach (var function in functions)
            {
                this.WriteHeading(function.Name, headingLevel + 1);

                var paragraph = this.CreateParagraph();
                var run = paragraph.AppendChild(new Run(new Text(function.Definition), new Text(function.ArgsString)));
                run.FormatCode();
            }
        }

        private void WriteHeading(string text, int headingLevel)
        {
            var heading = this.CreateParagraph();
            heading.AppendChild(new Run(new Text(text)));
            heading.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId() { Val = $"Heading{headingLevel}" });
        }

        private void WriteDescriptions(Descriptions descriptions)
        {
            this.WriteTextParagraph(descriptions.BriefDescription);
            foreach (var paragraph in descriptions.DetailedDescription)
            {
                this.WriteTextParagraph(paragraph);
            }
        }

        private void WriteTextParagraph(TextParagraph textParagraph)
        {
            // TODO: Type (warning, etc)

            var paragraph = this.CreateParagraph();
            foreach (var textRun in textParagraph)
            {
                switch (textRun)
                {
                    case TextRun t:
                        var run = paragraph.AppendChild(new Run(new Text(t.Text) { Space = SpaceProcessingModeValues.Preserve }));
                        run.RunProperties = new RunProperties()
                        {
                            Bold = t.Format.HasFlag(TextRunFormat.Bold) ? new Bold() : null,
                            Italic = t.Format.HasFlag(TextRunFormat.Italic) ? new Italic() : null,
                        };
                        if (t.Format.HasFlag(TextRunFormat.Monospace))
                        {
                            run.PrependChild(new RunFonts() { Ascii = "Consolas" });
                        }
                        break;
                }
            }
        }

        private Paragraph CreateParagraph()
        {
            var p = new Paragraph();
            this.paragraphs.Add(p);
            return p;
        }
    }
}
