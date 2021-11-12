﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Dox2Word.Logging;
using Dox2Word.Model;

namespace Dox2Word.Generator
{
    public class WordGenerator
    {
        public const string Placeholder = "<INSERT HERE>";
        private static readonly Logger logger = Logger.Instance;

        private readonly WordprocessingDocument doc;

        private readonly List<OpenXmlElement> bodyElements = new();

        private readonly ListManager listManager;

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
            var stylesPart = doc.MainDocumentPart.StyleDefinitionsPart;
            if (stylesPart == null)
            {
                stylesPart = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                new Styles().Save(stylesPart);
            }

            this.listManager = new ListManager(numberingPart);
            StyleManager.EnsureStyles(stylesPart);
        }

        public void Generate(Project project)
        {
            try
            {
                var body = this.doc.MainDocumentPart!.Document.Body!;

                var markerParagraph = body.ChildElements.OfType<Paragraph>()
                    .FirstOrDefault(x => x.InnerText == Placeholder);

                if (markerParagraph == null)
                    throw new GeneratorException($"Could not find placeholder text '{Placeholder}'");

                var headingBefore = markerParagraph.ElementsBefore().OfType<Paragraph>().FirstOrDefault();
                int headingLevel = 2;
                if (headingBefore?.ParagraphProperties?.ParagraphStyleId?.Val?.Value is { } styleId && styleId.StartsWith("Heading"))
                {
                    if (int.TryParse(styleId.Substring("Heading".Length), out int level))
                    {
                        headingLevel = level + 1;
                    }
                }

                foreach (var group in project.Groups)
                {
                    this.WriteGroup(group, headingLevel);
                }

                foreach (var paragraph in this.bodyElements)
                {
                    body.InsertBefore(paragraph, markerParagraph);
                }

                body.RemoveChild(markerParagraph);

                this.Validate();
            }
            catch (GeneratorException e)
            {
                logger.Error(e);
            }
        }

        private void Validate()
        {
            var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
            foreach (var error in validator.Validate(this.doc))
            {
                logger.Warning($"Validation failure. Description: {error.Description}; ErrorType: {error.ErrorType}; Node: {error.Node}; Path: {error.Path?.XPath}; Part: {error.Part?.Uri}");
            }
        }

        private void WriteGroup(Group group, int headingLevel)
        {
            logger.Info($"Writing group {group.Name}");
            this.WriteHeading(group.Name, headingLevel);

            this.WriteDescriptions(group.Descriptions);

            if (group.Files.Count > 0)
            {
                this.WriteMiniHeading("Files");
                var listParagraph = new ListParagraph(ListParagraphType.Bullet);
                listParagraph.Items.AddRange(group.Files.Select(x => new TextParagraph() { new TextRun(x) }));
                this.Append(this.CreateParagraph(listParagraph));
            }

            this.WriteGlobalVariables(group.GlobalVariables, headingLevel + 1);
            this.WriteMacros(group.Macros, headingLevel + 1);
            this.WriteEnums(group.Enums, headingLevel + 1);
            this.WriteClasses(group.Classes, headingLevel + 1);
            this.WriteFunctions(group.Functions, headingLevel + 1);

            foreach (var subGroup in group.SubGroups)
            {
                this.WriteGroup(subGroup, headingLevel + 1);
            }
        }

        private void WriteClasses(List<ClassDoc> classes, int headingLevel)
        {
            foreach (var cls in classes)
            {
                logger.Info($"Writing {cls.Type} {cls.Name}");

                string title = cls.Type switch
                {
                    ClassType.Struct => "Struct",
                    ClassType.Union => "Union",
                };

                this.WriteHeading($"{title} {cls.Name}", headingLevel);

                this.WriteDescriptions(cls.Descriptions);

                if (cls.Variables.Count > 0)
                {
                    this.WriteMiniHeading("Members");
                    var table = this.AppendChild(this.CreateTable().ApplyStyle(StyleManager.ParameterTableStyleId).AddColumns(2));
                    foreach (var variable in cls.Variables)
                    {
                        string? bitfield = variable.Bitfield == null
                            ? null
                            : $" :{variable.Bitfield}";
                        table.AppendRow($"{variable.Type} {variable.Name}{variable.ArgsString}{bitfield}", this.CreateDescriptions(variable.Descriptions));
                    }
                }
            }
        }

        private void WriteEnums(List<EnumDoc> enums, int headingLevel)
        {
            foreach (var @enum in enums)
            {
                logger.Info($"Writing Enum {@enum.Name}");

                this.WriteHeading($"Enum {@enum.Name}", headingLevel);

                this.WriteDescriptions(@enum.Descriptions);

                if (@enum.Values.Count > 0)
                {
                    this.WriteMiniHeading("Values");
                    var table = this.AppendChild(this.CreateTable().ApplyStyle(StyleManager.ParameterTableStyleId).AddColumns(2));
                    foreach (var value in @enum.Values)
                    {
                        table.AppendRow(value.Name, this.CreateDescriptions(value.Descriptions));
                    }
                }
            }
        }

        private void WriteGlobalVariables(List<VariableDoc> variables, int headingLevel)
        {
            foreach (var variable in variables)
            {
                logger.Info($"Writing variable {variable.Name}");

                this.WriteHeading($"Global variable {variable.Name}", headingLevel);

                this.WriteMiniHeading("Definition");
                var paragraph = this.AppendChild(new Paragraph().LeftAlign());
                var run = paragraph.AppendChild(new Run(new Text(variable.Definition ?? "")).ApplyStyle(StyleManager.CodeCharStyleId));
                if (variable.Initializer != null)
                {
                    run.Append(StringToElements(" " + variable.Initializer));
                }

                this.WriteDescriptions(variable.Descriptions);
            }
        }

        private void WriteMacros(List<MacroDoc> macros, int headingLevel)
        {
            foreach (var macro in macros)
            {
                logger.Info($"Writing macro {macro.Name}");

                this.WriteHeading($"Macro {macro.Name}", headingLevel);

                this.WriteMiniHeading($"Definition");

                // If it's a multi-line macro declaration, push it only a new line and add \ to the first line
                string initializer = macro.Initializer;
                if (initializer.StartsWith(" ") && initializer.Contains("\n"))
                {
                    initializer = "\\\n" + initializer;
                }
                string? paramsStr = macro.HasParameters
                    ? "(" + string.Join(", ", macro.Parameters.Select(x => x.Name)) + ")"
                    : null;

                this.AppendChild(new Paragraph(new Run(StringToElements($"#define {macro.Name}{paramsStr} {initializer}")).ApplyStyle(StyleManager.CodeCharStyleId)).LeftAlign());

                this.WriteDescriptions(macro.Descriptions);

                if (macro.Parameters.Count > 0 && macro.Parameters.Any(x => x.Description.Count > 0))
                {
                    this.WriteMiniHeading("Parameters");

                    var table = this.AppendChild(this.CreateTable().ApplyStyle(StyleManager.ParameterTableStyleId).AddColumns(2));
                    foreach (var parameter in macro.Parameters)
                    {
                        table.AppendRow(parameter.Name, this.CreateParagraph(parameter.Description));
                    }
                }

                this.WriteReturnDescription(macro.ReturnDescriptions);
            }
        }

        private void WriteFunctions(List<FunctionDoc> functions, int headingLevel)
        {
            foreach (var function in functions)
            {
                logger.Info($"Writing function {function.Name}");

                this.WriteHeading($"Function {function.Name}", headingLevel);

                this.WriteMiniHeading("Signature");
                var paragraph = this.AppendChild(new Paragraph().LeftAlign());
                var run = paragraph.AppendChild(new Run().ApplyStyle(StyleManager.CodeCharStyleId));
                run.AppendChild(new Text(function.Definition));

                if (function.Parameters.Count > 0)
                {
                    // It's safest to split the argsstring on ',' than to reconstruct each parameters.
                    // This is because the parameters don't contain info such as whether a '*' was next to the
                    // type or variable name, and it's hard to work this out after the fact.

                    run.AppendChild(new Text("("));
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
                    run.AppendChild(new Text(function.ArgsString));
                }

                this.WriteDescriptions(function.Descriptions);

                if (function.Parameters.Count > 0 && function.Parameters.Any(x => x.Description.Count > 0))
                {
                    this.WriteMiniHeading("Parameters");

                    bool hasInOut = function.Parameters.Any(x => x.Direction != ParameterDirection.None);
                    var table = this.AppendChild(this.CreateTable().ApplyStyle(StyleManager.ParameterTableStyleId).AddColumns(hasInOut ? 3 : 2));
                    foreach (var parameter in function.Parameters)
                    {
                        string? inOut = hasInOut
                            ? parameter.Direction switch
                            {
                                ParameterDirection.None => "",
                                ParameterDirection.In => "in",
                                ParameterDirection.Out => "out",
                                ParameterDirection.InOut => "in,out",
                            }
                            : null;
                        table.AppendRow(parameter.Name, this.CreateParagraph(parameter.Description), inOut);
                    }
                }

                this.WriteReturnDescription(function.ReturnDescriptions);
            }
        }

        private void WriteReturnDescription(ReturnDescriptions returnDescriptions)
        {
            if (returnDescriptions.Description.Count > 0)
            {
                this.WriteMiniHeading("Returns");
                this.Append(this.CreateParagraph(returnDescriptions.Description));
            }

            if (returnDescriptions.Values.Count > 0)
            {
                this.WriteMiniHeading("Return values");
                var table = this.AppendChild(this.CreateTable().ApplyStyle(StyleManager.ParameterTableStyleId).AddColumns(2));
                foreach (var returnValue in returnDescriptions.Values)
                {
                    table.AppendRow(returnValue.Name, this.CreateParagraph(returnValue.Description));
                }
            }
        }

        private void WriteHeading(string text, int headingLevel)
        {
            var paragraph = this.AppendChild(new Paragraph());
            var run = paragraph.AppendChild(new Run(new Text(text)));
            paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = $"Heading{headingLevel}" };
            // The style might not have enough spacing, so force this
            paragraph.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines()
            {
                Before = $"{8 * 20}",
            };

            // If the style specifies all-caps or small-caps, undo this
            run.RunProperties ??= new RunProperties();
            run.RunProperties.SmallCaps = new SmallCaps() { Val = false };
            run.RunProperties.Caps = new Caps() { Val = false };
        }

        private void WriteMiniHeading(string text)
        {
            this.AppendChild(new Paragraph(new Run(new Text(text)))).ApplyStyle(StyleManager.MiniHeadingStyleId);
        }

        private void WriteDescriptions(Descriptions descriptions)
        {
            if (descriptions.HasDescriptions)
            { 
                this.WriteMiniHeading("Description");
                this.Append(this.CreateDescriptions(descriptions));
            }
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

        private IEnumerable<OpenXmlElement> CreateParagraph(IParagraph inputParagraph)
        {
            switch (inputParagraph)
            {
                case TextParagraph textParagraph:
                    var paragraph = this.CreateTextParagraph(textParagraph);
                    if (textParagraph.Type == ParagraphType.Warning)
                    {
                        paragraph.ApplyStyle(StyleManager.WarningStyleId);
                    }
                    yield return paragraph;
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

                case TableDoc table:
                    yield return this.CreateTable(table);
                    break;

                default:
                    throw new GeneratorException($"Unexpected paragraph type {inputParagraph} ({inputParagraph.GetType()})");
            }
        }

        private Paragraph CreateTextParagraph(TextParagraph textParagraph)
        {
            var paragraph = new Paragraph();
            if (textParagraph.Alignment != ParagraphAlignment.Default)
            {
                var justification = textParagraph.Alignment switch
                {
                    ParagraphAlignment.Left => JustificationValues.Left,
                    ParagraphAlignment.Center => JustificationValues.Center,
                    ParagraphAlignment.Right => JustificationValues.Right,
                    ParagraphAlignment.Default => throw new Exception("Not possible"),
                };

                paragraph.ParagraphProperties ??= new ParagraphProperties();
                paragraph.ParagraphProperties.Justification = new Justification() { Val = justification };
            }

            foreach (var textRun in textParagraph)
            {
                var run = paragraph.AppendChild(new Run());
                var text = run.AppendChild(new Text(textRun.Text));
                if (textRun.Text.StartsWith(" ") || textRun.Text.EndsWith(" "))
                {
                    text.Space = SpaceProcessingModeValues.Preserve;
                }
                run.RunProperties = new RunProperties()
                {
                    Bold = textRun.Format.HasFlag(TextRunFormat.Bold) ? new Bold() : null,
                    Italic = textRun.Format.HasFlag(TextRunFormat.Italic) ? new Italic() : null,
                };
                if (textRun.Format.HasFlag(TextRunFormat.Monospace))
                {
                    run.RunProperties.RunFonts = new RunFonts() { Ascii = "Consolas" };
                }
            }
            return paragraph;
        }

        private IEnumerable<OpenXmlElement> CreateListParagraph(ListParagraph listParagraph, int level = 0)
        {
            int numberId = this.listManager.CreateList(listParagraph.Type);

            foreach (var listItem in listParagraph.Items)
            {
                var paragraphProperties = new ParagraphProperties()
                {
                    ParagraphStyleId = new ParagraphStyleId() { Val = StyleManager.ListParagraphStyleId },
                    NumberingProperties = new NumberingProperties()
                    {
                        NumberingLevelReference = new NumberingLevelReference() { Val = level },
                        NumberingId = new NumberingId() { Val = numberId },
                    },
                };

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
            var paragraph = new Paragraph().ApplyStyle(StyleManager.CodeStyleId).LeftAlign();

            var run = paragraph.AppendChild(new Run());
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

        private OpenXmlElement CreateTable(TableDoc tableDoc)
        {
            var table = this.CreateTable();
            var tableProperties = table.Elements<TableProperties>().FirstOrDefault() ??
                table.AppendChild(new TableProperties());
            tableProperties.TableStyle = new TableStyle() { Val = StyleManager.TableStyleId };
            tableProperties.TableLook = new TableLook()
            {
                FirstRow = tableDoc.FirstRowHeader,
                LastRow = false,
                FirstColumn = tableDoc.FirstColumnHeader,
                LastColumn = false,
            };
            table.AddColumns(tableDoc.NumColumns);

            foreach (var rowDoc in tableDoc.Rows)
            {
                var row = table.AppendChild(new TableRow());
                foreach (var cellDoc in rowDoc.Cells)
                {
                    var cell = row.AppendChild(new TableCell());
                    var valueList = cellDoc.Paragraphs.SelectMany(x => this.CreateParagraph(x)).ToList();
                    cell.Append(valueList);
                    for (int i = 0; i < valueList.Count; i++)
                    {
                        valueList[i].FormatTableCellElement(after: i == valueList.Count - 1);
                    }
                }
            }

            return table;
        }

        private static IEnumerable<OpenXmlElement> StringToElements(string? input)
        {
            if (input == null)
                yield break;

            string[] lines = input.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                if (i > 0)
                {
                    yield return new Break();
                }
                var text = new Text(lines[i]);
                if (lines[i].StartsWith(" "))
                {
                    text.Space = SpaceProcessingModeValues.Preserve;
                }
                yield return text;
            }
        }

        private T AppendChild<T>(T child) where T : OpenXmlElement
        {
            this.bodyElements.Add(child);
            return child;
        }

        private Table CreateTable()
        {
            var table = new Table();

            // Add spacing after
            //var paragraph = this.AppendChild(new Paragraph());
            //paragraph.ParagraphProperties ??= new ParagraphProperties();
            //paragraph.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines()
            //{
            //    Before = "0",
            //    After = "0",
            //};

            //var run = paragraph.AppendChild(new Run());
            //run.RunProperties ??= new RunProperties();
            //run.RunProperties.FontSize = new FontSize() { Val = "10" };

            return table;
        }

        private void Append(IEnumerable<OpenXmlElement> children)
        {
            this.bodyElements.AddRange(children);
        }
    }
}
