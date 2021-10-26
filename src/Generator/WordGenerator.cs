﻿using System;
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

        private readonly WordprocessingDocument doc;

        private readonly List<OpenXmlElement> bodyElements = new();

        private readonly ListStyles listStyles;

        public static void Generate(Stream stream, Project project)
        {
            using var doc = WordprocessingDocument.Open(stream, isEditable: true);
            new WordGenerator(doc).Generate(project);
        }

        private WordGenerator(WordprocessingDocument doc)
        {
            this.doc = doc;
            var numberingPart = doc.MainDocumentPart!.NumberingDefinitionsPart!;
            if (numberingPart == null)
            {
                numberingPart = doc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>("NumberingDefinitionsPart001");
                new Numbering().Save(numberingPart);
            }

            this.listStyles = new ListStyles(numberingPart);
        }

        public void Generate(Project project)
        {
            var body = this.doc.MainDocumentPart!.Document.Body!;

            foreach (var group in project.Groups)
            {
                this.WriteGroup(group, 2);
            }

            foreach (var paragraph in this.bodyElements)
            {
                body.AppendChild(paragraph);
            }
        }

        private void WriteGroup(Group group, int headingLevel)
        {
            this.WriteHeading(group.Name, headingLevel);

            this.Append(this.CreateDescriptions(group.Descriptions));

            this.WriteClasses(group.Classes, headingLevel + 1);
            this.WriteEnums(group.Enums, headingLevel + 1);
            this.WriteGlobalVariables(group.GlobalVariables, headingLevel + 1);
            this.WriteFunctions(group.Functions, headingLevel + 1);

            foreach (var subGroup in group.SubGroups)
            {
                this.WriteGroup(subGroup, headingLevel + 1);
            }
        }

        private void WriteClasses(List<ClassDoc> classes, int headingLevel)
        {
            if (classes.Count == 0)
                return;

            this.WriteHeading("Structs", headingLevel);

            foreach (var cls in classes)
            {
                this.WriteHeading(cls.Name, headingLevel + 1);

                this.Append(this.CreateDescriptions(cls.Descriptions));

                if (cls.Variables.Count > 0)
                {
                    var table = this.AppendChild(new Table().AddBorders());
                    foreach (var variable in cls.Variables)
                    {
                        var row = table.AppendChild(new TableRow());
                        var nameCell = row.AppendChild(CreateTableCell());
                        nameCell.Append(new Paragraph(new Run(new Text($"{variable.Type} {variable.Name}")).FormatCode()));
                        var descriptionCell = row.AppendChild(CreateTableCell()); ;
                        descriptionCell.Append(this.CreateDescriptions(variable.Descriptions));
                    }
                }
            }
        }

        private void WriteEnums(List<EnumDoc> enums, int headingLevel)
        {
            if (enums.Count == 0)
                return;

            this.WriteHeading("Enums", headingLevel);

            foreach (var @enum in enums)
            {
                this.WriteHeading(@enum.Name, headingLevel + 1);

                this.Append(this.CreateDescriptions(@enum.Descriptions));

                if (@enum.Values.Count > 0)
                {
                    var table = this.AppendChild(new Table().AddBorders());
                    foreach (var value in @enum.Values)
                    {
                        var row = table.AppendChild(new TableRow());
                        var nameCell = row.AppendChild(CreateTableCell());
                        nameCell.Append(new Paragraph(new Run(new Text(value.Name)).FormatCode()));
                        var descriptionCell = row.AppendChild(CreateTableCell());
                        descriptionCell.Append(this.CreateDescriptions(value.Descriptions));
                    }
                }
            }
        }

        private void WriteGlobalVariables(List<VariableDoc> variables, int headingLevel)
        {
            if (variables.Count == 0)
                return;

            this.WriteHeading("Global Variables", headingLevel);

            foreach (var variable in variables)
            {
                this.WriteHeading(variable.Name, headingLevel + 1);

                var paragraph = this.AppendChild(new Paragraph());
                var run = paragraph.AppendChild(new Run(new Text(variable.Definition)));
                run.FormatCode();

                this.Append(this.CreateDescriptions(variable.Descriptions));
            }
        }

        private void WriteFunctions(List<FunctionDoc> functions, int headingLevel)
        {
            if (functions.Count == 0)
                return;

            this.WriteHeading("Functions", headingLevel);

            foreach (var function in functions)
            {
                this.WriteHeading(function.Name, headingLevel + 1);

                this.WriteMiniHeading("Signature");
                var paragraph = this.AppendChild(new Paragraph());
                var run = paragraph.AppendChild(new Run().FormatCode());
                run.AppendChild(new Text(function.Definition + "("));

                if (function.Parameters.Count > 0)
                {
                    // It's safest to split the argsstring on ',' than to reconstruct each parameters.
                    // This is because the parameters don't contain info such as whether a '*' was next to the
                    // type or variable name, and it's hard to work this out after the fact.

                    string[] parameters = function.ArgsString.TrimStart('(').TrimEnd(')').Split(',');
                    for (int i = 0; i < parameters.Length; i++)
                    {
                        string parameter = parameters[i];
                        string comma = i == parameters.Length - 1 ? "" : ",";
                        run.AppendChild(new Break());
                        run.AppendChild(new Text($"    {parameter.Trim()}{comma}")
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        });
                    }
                    run.AppendChild(new Text(")"));
                }
                else
                {
                    run.AppendChild(new Text("void)"));
                }

                this.WriteMiniHeading("Description");
                this.Append(this.CreateDescriptions(function.Descriptions));

                if (function.Parameters.Count > 0)
                {
                    this.WriteMiniHeading("Parameters");
                    var table = this.AppendChild(new Table().AddBorders());
                    foreach (var parameter in function.Parameters)
                    {
                        var row = table.AppendChild(new TableRow());
                        var nameCell = row.AppendChild(CreateTableCell());
                        nameCell.AppendChild(new Paragraph(new Run(new Text(parameter.Name)).FormatCode()));
                        var descriptionCell = row.AppendChild(CreateTableCell());
                        descriptionCell.Append(this.CreateParagraph(parameter.Description));
                    }
                }

                if (function.ReturnDescription.Count > 0)
                {
                    this.WriteMiniHeading("Returns");
                    this.Append(this.CreateParagraph(function.ReturnDescription));
                }

                if (function.ReturnValues.Count > 0)
                {
                    this.WriteMiniHeading("Return values");
                    var table = this.AppendChild(new Table().AddBorders());
                    foreach (var returnValue in function.ReturnValues)
                    {
                        var row = table.AppendChild(new TableRow());
                        var valueCell = row.AppendChild(CreateTableCell());
                        valueCell.AppendChild(new Paragraph(new Run(new Text(returnValue.Name)).FormatCode()));
                        var descriptionCell = row.AppendChild(CreateTableCell());
                        descriptionCell.Append(this.CreateParagraph(returnValue.Description));
                    }
                }
            }
        }

        private void WriteHeading(string text, int headingLevel)
        {
            var heading = this.AppendChild(new Paragraph());
            heading.AppendChild(new Run(new Text(text)));
            heading.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId() { Val = $"Heading{headingLevel}" });
        }

        private void WriteMiniHeading(string text)
        {
            var heading = this.AppendChild(new Paragraph());
            heading.AppendChild(new Run(new Text(text)).FormatMiniHeading());
        }

        private IEnumerable<OpenXmlElement> CreateDescriptions(Descriptions descriptions)
        {
            foreach (var p in this.CreateParagraph(descriptions.BriefDescription))
            {
                yield return p;
            }
            foreach (var paragraph in descriptions.DetailedDescription)
            {
                foreach (var p in this.CreateParagraph(paragraph))
                {
                    yield return p;
                }
            }
        }

        private IEnumerable<Paragraph> CreateParagraph(IParagraph inputParagraph)
        {
            // TODO: Type (warning, etc)

            switch (inputParagraph)
            {
                case TextParagraph textParagraph:
                    yield return this.CreateTextParagraph(textParagraph);
                    break;

                case ListParagraph listParagraph:
                    foreach (var p in this.CreateListParagraph(listParagraph))
                    {
                        yield return p;
                    }
                    break;

                case CodeParagraph codeParagraph:
                    yield return this.CreateCodeParagraph(codeParagraph);
                    break;

                default:
                    throw new GeneratorException($"Unexpected paragraph type {inputParagraph} ({inputParagraph.GetType()})");
            }
        }

        private Paragraph CreateTextParagraph(TextParagraph textParagraph)
        {
            var paragraph = new Paragraph();
            foreach (var textRun in textParagraph)
            {
                var run = paragraph.AppendChild(new Run(new Text(textRun.Text) { Space = SpaceProcessingModeValues.Preserve }));
                run.RunProperties = new RunProperties()
                {
                    Bold = textRun.Format.HasFlag(TextRunFormat.Bold) ? new Bold() : null,
                    Italic = textRun.Format.HasFlag(TextRunFormat.Italic) ? new Italic() : null,
                };
                if (textRun.Format.HasFlag(TextRunFormat.Monospace))
                {
                    run.PrependChild(new RunFonts() { Ascii = "Consolas" });
                }
            }
            return paragraph;
        }

        private IEnumerable<Paragraph> CreateListParagraph(ListParagraph listParagraph, int level = 0)
        {
            int numberId = this.listStyles.CreateList(listParagraph.Type);

            foreach (var listItem in listParagraph.Items)
            {
                var paragraphProperties = new ParagraphProperties();
                paragraphProperties.AppendChild(new NumberingProperties(
                    new NumberingLevelReference() { Val = level },
                    new NumberingId() { Val = numberId }));
                paragraphProperties.AppendChild(new SpacingBetweenLines() { After = "0" });  // Get rid of space between bullets

                // Single paragraphs are a list item. Child lists are rendered as a child list.
                // Everything else doesn't have the list style applied to it, and the list will continue after it
                switch (listItem)
                {
                    case TextParagraph textParagraph:
                        var paragraph = this.CreateTextParagraph(textParagraph);
                        paragraph.ParagraphProperties = paragraphProperties;
                        yield return paragraph;
                        break;

                    case ListParagraph childListParagraph:
                        foreach (var p in this.CreateListParagraph(childListParagraph, level + 1))
                        {
                            yield return p;
                        }
                        break;

                    default:
                        foreach (var p in this.CreateParagraph(listItem))
                        {
                            yield return p;
                        }
                        break;
                }
            }
        }

        private Paragraph CreateCodeParagraph(CodeParagraph codeParagraph)
        {
            var paragraph = new Paragraph();
            var paragraphProperties = paragraph.AppendChild(new ParagraphProperties());
            paragraphProperties.AppendChild(new ParagraphBorders(
                new TopBorder() { Val = BorderValues.Single },
                new RightBorder() { Val = BorderValues.Single },
                new BottomBorder() { Val = BorderValues.Single },
                new LeftBorder() { Val = BorderValues.Single }));

            var run = paragraph.AppendChild(new Run().FormatCode());
            for (int i = 0; i < codeParagraph.Lines.Count; i++)
            {
                if (i > 0)
                {
                    run.AppendChild(new Break());
                }
                run.AppendChild(new Text(codeParagraph.Lines[i]) { Space = SpaceProcessingModeValues.Preserve });
            }
            return paragraph;
        }

        private T AppendChild<T>(T child) where T : OpenXmlElement
        {
            this.bodyElements.Add(child);
            return child;
        }

        private void Append(IEnumerable<OpenXmlElement> children)
        {
            this.bodyElements.AddRange(children);
        }

        private static TableCell CreateTableCell()
        {
            var cell = new TableCell();
            //cell.AppendChild(new TableCellProperties(
            //    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
            //));
            return cell;
        }
    }
}
