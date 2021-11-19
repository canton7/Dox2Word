using System;
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
        public const string Placeholder = "MODULES";
        private static readonly Logger logger = Logger.Instance;

        private readonly WordprocessingDocument doc;

        private readonly List<OpenXmlElement> bodyElements = new();

        private readonly ListManager listManager;
        private readonly ImageManager imageManager;
        private readonly BookmarkManager bookmarkManager;
        private readonly DotRenderer dotRenderer = new();

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
            this.imageManager = new ImageManager(doc.MainDocumentPart);
            this.bookmarkManager = new BookmarkManager(doc.MainDocumentPart.Document.Body!);
            StyleManager.EnsureStyles(stylesPart);
        }

        public void Generate(Project project)
        {
            try
            {
                var body = this.doc.MainDocumentPart!.Document.Body!;

                var markerParagraph = PlaceholderHelper.FindParagraphPlaceholder(body, Placeholder);
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

                foreach (var group in project.RootGroups)
                {
                    this.WriteGroup(group, headingLevel);
                }

                foreach (var paragraph in this.bodyElements)
                {
                    body.InsertBefore(paragraph, markerParagraph);
                }

                body.RemoveChild(markerParagraph);

                this.SubstituteOptionPlaceholders(body, project);

                this.Validate();
            }
            catch (GeneratorException e)
            {
                logger.Error(e);
            }
        }

        private void SubstituteOptionPlaceholders(Body body, Project project)
        {
            var options = project.Options.ToDictionary(x => x.Id, x => string.Join(", ", x.Values));
            PlaceholderHelper.Replace(body, options);
        }

        private void Validate()
        {
            var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
            foreach (var error in validator.Validate(this.doc))
            {
                logger.Warning($"Validation failure. Description: {error.Description}; ErrorType: {error.ErrorType}; Node: {error.Node}; Path: {error.Path?.XPath}; Part: {error.Part?.Uri}");
            }
            this.bookmarkManager.Validate();
        }

        private void WriteGroup(Group group, int headingLevel)
        {
            logger.Info($"Writing group {group.Name}");
            this.WriteHeading(null, group.Name, group.Id, headingLevel);

            this.WriteDescriptions(group.Descriptions);

            if (group.Files.Count > 0)
            {
                this.WriteMiniHeading("Files");
                var filesList = new ListParagraph(ListParagraphType.Bullet);
                filesList.Items.AddRange(group.Files.Select(x => new TextParagraph() { new TextRun(x.Name) }));
                this.Append(this.CreateParagraph(filesList));

                if (group.IncludedGroups.Count > 0 || group.IncludingGroups.Count > 0)
                {
                    this.WriteMiniHeading("Unit References");
                    if (group.IncludedGroups.Count > 0)
                    {
                        this.AppendChild(StringToParagraph("This unit references the following units:"));
                        var groupsList = new ListParagraph(ListParagraphType.Bullet);
                        groupsList.Items.AddRange(group.IncludedGroups.Select(x => new TextParagraph() { new TextRun(x.Name, TextRunFormat.Monospace, x.Id) }));
                        this.Append(this.CreateParagraph(groupsList));
                    }
                    if (group.IncludingGroups.Count > 0)
                    {
                        this.AppendChild(StringToParagraph("This unit is referenced by the following units:"));
                        var groupsList = new ListParagraph(ListParagraphType.Bullet);
                        groupsList.Items.AddRange(group.IncludingGroups.Select(x => new TextParagraph() { new TextRun(x.Name, TextRunFormat.Monospace, x.Id) }));
                        this.Append(this.CreateParagraph(groupsList));
                    }
                }
            }

            this.WriteTypedefs(group.Typedefs, headingLevel + 1);
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

                this.WriteHeading(title, cls.Name, cls.Id, headingLevel);

                this.WriteDescriptions(cls.Descriptions);

                if (cls.Variables.Count > 0)
                {
                    this.WriteMiniHeading("Members");
                    var table = this.AppendChild(this.CreateTable(StyleManager.ParameterTableStyleId).AddColumns(2));
                    foreach (var variable in cls.Variables)
                    {
                        string? bitfield = variable.Bitfield == null
                            ? null
                            : $" :{variable.Bitfield}";
                        table.AppendRow(this.TextRunsToRuns(variable.Type).Append(new Run(new Text($" {variable.Name}{variable.ArgsString}{bitfield}").PreserveSpace())),
                            this.bookmarkManager.CreateBookmark(variable.Id), this.CreateDescriptions(variable.Descriptions));
                    }
                }
            }
        }

        private void WriteEnums(List<EnumDoc> enums, int headingLevel)
        {
            foreach (var @enum in enums)
            {
                logger.Info($"Writing Enum {@enum.Name}");

                this.WriteHeading("Enum", @enum.Name, @enum.Id, headingLevel);

                this.WriteDescriptions(@enum.Descriptions);

                if (@enum.Values.Count > 0)
                {
                    this.WriteMiniHeading("Values");
                    var table = this.AppendChild(this.CreateTable(StyleManager.ParameterTableStyleId).AddColumns(2));
                    foreach (var value in @enum.Values)
                    {
                        table.AppendRow(value.Name, this.bookmarkManager.CreateBookmark(value.Id), this.CreateDescriptions(value.Descriptions));
                    }
                }
            }
        }

        private void WriteTypedefs(List<TypedefDoc> typedefs, int headingLevel)
        {
            foreach (var typedef in typedefs)
            {
                logger.Info($"Writing typedef {typedef.Name}");

                this.WriteHeading("Typedef", typedef.Name, typedef.Id, headingLevel);

                this.WriteMiniHeading("Definition");
                var paragraph = this.AppendChild(new Paragraph().LeftAlign());
                paragraph.AppendChild(new Run(new Text("typedef ").PreserveSpace()).ApplyStyle(StyleManager.CodeCharStyleId));
                paragraph.Append(this.TextRunsToRuns(typedef.Type).ApplyStyle(StyleManager.CodeCharStyleId));
                paragraph.AppendChild(new Run(new Text(" " + typedef.Name).PreserveSpace()).ApplyStyle(StyleManager.CodeCharStyleId));

                this.WriteDescriptions(typedef.Descriptions);
            }
        }

        private void WriteGlobalVariables(List<VariableDoc> variables, int headingLevel)
        {
            foreach (var variable in variables)
            {
                logger.Info($"Writing variable {variable.Name}");

                this.WriteHeading("Global variable", variable.Name!, variable.Id, headingLevel);

                this.WriteMiniHeading("Definition");
                var paragraph = this.AppendChild(new Paragraph().LeftAlign());
                paragraph.Append(this.TextRunsToRuns(variable.Type).ApplyStyle(StyleManager.CodeCharStyleId));
                paragraph.Append(new Run(new Text($" {variable.Name}").PreserveSpace()).ApplyStyle(StyleManager.CodeCharStyleId));
                if (variable.Initializer.Count > 0)
                {
                    paragraph.AppendChild(new Run(new Text(" ").PreserveSpace()).ApplyStyle(StyleManager.CodeCharStyleId));
                    paragraph.Append(this.TextRunsToRuns(variable.Initializer).ApplyStyle(StyleManager.CodeCharStyleId));
                }

                this.WriteDescriptions(variable.Descriptions);
            }
        }

        private void WriteMacros(List<MacroDoc> macros, int headingLevel)
        {
            foreach (var macro in macros)
            {
                logger.Info($"Writing macro {macro.Name}");

                this.WriteHeading("Macro", macro.Name, macro.Id, headingLevel);

                this.WriteMiniHeading($"Definition");

                string? paramsStr = macro.HasParameters
                    ? "(" + string.Join(", ", macro.Parameters.Select(x => x.Name)) + ")"
                    : null;
                var paragraph = this.AppendChild(new Paragraph(new Run(new Text($"#define {macro.Name}{paramsStr} ")
                    .PreserveSpace()).ApplyStyle(StyleManager.CodeCharStyleId)).LeftAlign());

                // If it's a multi-line macro declaration, push it onto a new line and add \ to the first line
                if (macro.Initializer.Count > 0 && macro.Initializer[0].Text.StartsWith(" ") &&
                    macro.Initializer.Any(x => x.Text.Contains("\n")))
                {
                    paragraph.AppendChild(new Run(new Text("\\"), new Break()));
                }
                paragraph.Append(this.TextRunsToRuns(macro.Initializer).ApplyStyle(StyleManager.CodeCharStyleId));

                this.WriteDescriptions(macro.Descriptions);

                if (macro.Parameters.Count > 0 && macro.Parameters.Any(x => !x.Description.IsEmpty))
                {
                    this.WriteMiniHeading("Parameters");

                    var table = this.AppendChild(this.CreateTable(StyleManager.ParameterTableStyleId).AddColumns(2));
                    foreach (var parameter in macro.Parameters)
                    {
                        table.AppendRow(parameter.Name, null, this.CreateParagraph(parameter.Description));
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

                this.WriteHeading("Function", function.Name, function.Id, headingLevel);

                this.WriteMiniHeading("Signature");
                var paragraph = this.AppendChild(new Paragraph().LeftAlign());
                Run NewRun(OpenXmlElement e) => paragraph.AppendChild(new Run(e).ApplyStyle(StyleManager.CodeCharStyleId));

                paragraph.Append(this.TextRunsToRuns(function.ReturnType).ApplyStyle(StyleManager.CodeCharStyleId));
                NewRun(new Text($" {function.Name}").PreserveSpace());

                if (function.Parameters.Count > 0)
                {
                    NewRun(new Text("("));
                    for (int i = 0; i < function.Parameters.Count;i++)
                    {
                        string comma = i == function.Parameters.Count - 1 ? "" : ",";
                        var run = NewRun(new Break());
                        run.AppendChild(new Text("    ").PreserveSpace());
                        paragraph.Append(this.TextRunsToRuns(function.Parameters[i].Type).ApplyStyle(StyleManager.CodeCharStyleId));
                        string space = paragraph.InnerText.EndsWith(" *") ? "" : " ";
                        NewRun(new Text($"{space}{function.Parameters[i].Name}{comma}").PreserveSpace());
                    }

                    NewRun(new Text(")"));
                }
                else
                {
                    NewRun(new Text(function.ArgsString));
                }

                this.WriteDescriptions(function.Descriptions);

                if (function.Parameters.Count > 0 && function.Parameters.Any(x => !x.Description.IsEmpty))
                {
                    this.WriteMiniHeading("Parameters");

                    bool hasInOut = function.Parameters.Any(x => x.Direction != ParameterDirection.None);
                    var table = this.AppendChild(this.CreateTable(StyleManager.ParameterTableStyleId).AddColumns(hasInOut ? 3 : 2));
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
                        table.AppendRow(parameter.Name, null, this.CreateParagraph(parameter.Description), inOut);
                    }
                }

                this.WriteReturnDescription(function.ReturnDescriptions);
            }
        }

        private void WriteReturnDescription(ReturnDescriptions returnDescriptions)
        {
            if (!returnDescriptions.Description.IsEmpty)
            {
                this.WriteMiniHeading("Returns");
                this.Append(this.CreateParagraph(returnDescriptions.Description));
            }

            if (returnDescriptions.Values.Count > 0)
            {
                this.WriteMiniHeading("Return values");
                var table = this.AppendChild(this.CreateTable(StyleManager.ParameterTableStyleId).AddColumns(2));
                foreach (var returnValue in returnDescriptions.Values)
                {
                    table.AppendRow(returnValue.Name, null, this.CreateParagraph(returnValue.Description));
                }
            }
        }

        private void WriteHeading(string? prefix, string text, string id, int headingLevel)
        {
            var paragraph = this.AppendChild(new Paragraph());

            if (prefix != null)
            {
                paragraph.Append(new Run(new Text(prefix + " ").PreserveSpace()));
            }

            var run = new Run(new Text(text));
            var (bookmarkStart, bookmarkEnd) = this.bookmarkManager.CreateBookmark(id);
            paragraph.Append(bookmarkStart, run, bookmarkEnd);

            // The style might not have enough spacing, so force this
            paragraph.ApplyStyle($"Heading{headingLevel}")
                .WithProperties(x => x.SpacingBetweenLines = new SpacingBetweenLines() { Before = $"{8 * 20}" });

            // If the style specifies all-caps or small-caps, undo this
            run.WithProperties(x =>
            {
                x.SmallCaps = new SmallCaps() { Val = false };
                x.Caps = new Caps() { Val = false };
            });
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
                    if (textParagraph.Type == TextParagraphType.Warning)
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

                case DotParagraph dotParagraph:
                    var para = this.CreateDotParagraph(dotParagraph);
                    if (para != null)
                    {
                        yield return para;
                    }
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
            if (textParagraph.Alignment != TextParagraphAlignment.Default)
            {
                var justification = textParagraph.Alignment switch
                {
                    TextParagraphAlignment.Left => JustificationValues.Left,
                    TextParagraphAlignment.Center => JustificationValues.Center,
                    TextParagraphAlignment.Right => JustificationValues.Right,
                    TextParagraphAlignment.Default => throw new Exception("Not possible"),
                };

                paragraph.WithProperties(x => x.Justification = new Justification() { Val = justification });
            }

            paragraph.Append(this.TextRunsToRuns(textParagraph));
            
            return paragraph;
        }

        private IEnumerable<Run> TextRunsToRuns(IEnumerable<TextRun> textRuns)
        {
            foreach (var textRun in textRuns)
            {
                var run = new Run(StringToElements(textRun.Text));

                run.WithProperties(x =>
                {
                    x.Bold = textRun.Format.HasFlag(TextRunFormat.Bold) ? new Bold() : null;
                    x.Italic = textRun.Format.HasFlag(TextRunFormat.Italic) ? new Italic() : null;
                });
                if (textRun.Format.HasFlag(TextRunFormat.Monospace))
                {
                    run.WithProperties(x =>
                    {
                        x.FontSize = new FontSize() { Val = "20" };
                        x.RunFonts = new RunFonts() { Ascii = "Consolas" };
                    });
                }

                if (textRun.ReferenceId == null)
                {
                    yield return run;
                }
                else
                {
                    foreach (var x in this.bookmarkManager.CreateLink(textRun.ReferenceId, run))
                    {
                        yield return x;
                    }
                }
            }
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
                run.AppendChild(new Text(codeParagraph.Lines[i]).PreserveSpace());
            }

            return paragraph;
        }

        private OpenXmlElement? CreateDotParagraph(DotParagraph dotParagraph)
        {
            byte[]? output = this.dotRenderer.TryRender(dotParagraph.Contents);
            if (output != null)
            {
                return new Paragraph(new Run(this.imageManager.CreateImage(output, ImagePartType.Png)));
            }
            return null;
        }

        private OpenXmlElement CreateTable(TableDoc tableDoc)
        {
            var table = new Table();
            table.WithProperties(x =>
            {
                x.TableStyle = new TableStyle() { Val = StyleManager.TableStyleId };
                x.TableLook = new TableLook()
                {
                    FirstRow = tableDoc.FirstRowHeader,
                    LastRow = false,
                    FirstColumn = tableDoc.FirstColumnHeader,
                    LastColumn = false,
                };
            });
            table.AddColumns(tableDoc.NumColumns);

            // Column index -> how many rowspans are left
            int[] colIndexToRowSpan = new int[tableDoc.NumColumns];

            foreach (var rowDoc in tableDoc.Rows)
            {
                var row = table.AppendChild(new TableRow());
                for (int i = 0, cellDocIndex = 0; i < tableDoc.NumColumns; i++, cellDocIndex++)
                {
                    var cell = row.AppendChild(new TableCell());

                    // If we're currently spanning this column there won't be an entry in rowDoc.Cells, and we'll
                    // have to synthesise it
                    if (colIndexToRowSpan[i] > 0)
                    {
                        cell.WithProperties(x => x.VerticalMerge = new VerticalMerge());
                        cell.Append(new Paragraph());
                        colIndexToRowSpan[i]--;
                        cellDocIndex--;
                    }
                    else
                    {
                        var cellDoc = rowDoc.Cells[cellDocIndex];

                        if (cellDoc.ColSpan > 1)
                        {
                            cell.WithProperties(x => x.GridSpan = new GridSpan() { Val = cellDoc.ColSpan });
                            i++;
                        }
                        if (cellDoc.RowSpan > 1)
                        {
                            cell.WithProperties(x => x.VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart });
                            colIndexToRowSpan[i] = cellDoc.RowSpan - 1;
                        }

                        var valueList = cellDoc.Paragraphs.SelectMany(x => this.CreateParagraph(x)).ToList();
                        cell.Append(valueList);
                        for (int j = 0; j < valueList.Count; j++)
                        {
                            valueList[j].FormatTableCellElement(after: j == valueList.Count - 1);
                        }
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
                if (lines[i].StartsWith(" ") || lines[i].EndsWith(" "))
                {
                    text.PreserveSpace();
                }
                yield return text;
            }
        }

        private T AppendChild<T>(T child) where T : OpenXmlElement
        {
            this.bodyElements.Add(child);
            return child;
        }

        private static Paragraph StringToParagraph(string text) => new Paragraph(new Run(new Text(text)));

        private Table CreateTable(string styleId)
        {
            return new Table().ApplyStyle(styleId)
                .WithProperties(x => x.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Auto, Width = "0" });
        }

        private void Append(IEnumerable<OpenXmlElement> children)
        {
            this.bodyElements.AddRange(children);
        }
    }
}
