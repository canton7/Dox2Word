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
        private readonly Options options;
        private readonly List<OpenXmlElement> bodyElements = new();

        private readonly ListManager listManager;
        private readonly ImageManager imageManager;
        private readonly BookmarkManager bookmarkManager;
        private readonly StyleManager styleManager;
        private readonly DotRenderer dotRenderer = new();

        public static void Generate(Stream stream, Project project, Options options)
        {
            using var doc = WordprocessingDocument.Open(stream, isEditable: true);
            new WordGenerator(doc, options).Generate(project);
        }

        private WordGenerator(WordprocessingDocument doc, Options options)
        {
            this.doc = doc;
            this.options = options;

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

            this.imageManager = new ImageManager(doc.MainDocumentPart);
            this.bookmarkManager = new BookmarkManager(doc.MainDocumentPart.Document.Body!);
            this.styleManager = new StyleManager(stylesPart);
            this.styleManager.EnsureStyles();
            this.listManager = new ListManager(numberingPart, this.styleManager);
            this.listManager.EnsureNumbers();
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

                this.SubstitutePlaceholders(body, project);

                this.Validate();
            }
            catch (GeneratorException e)
            {
                logger.Error(e);
            }
        }

        private void SubstitutePlaceholders(Body body, Project project)
        {
            var placeholders = project.Options.ToDictionary(x => $"DOXYFILE_{x.Id}", x => string.Join(", ", x.Values));
            foreach (var kvp in this.options.Placeholders)
            {
                placeholders[kvp.Key] = kvp.Value;
            }
            PlaceholderHelper.Replace(body, placeholders);
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
            logger.Debug($"Writing group {group.Name}");
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
                        groupsList.Items.AddRange(group.IncludedGroups.Select(x => new TextParagraph() { new TextRun(x.Name, new(TextRunFormat.Monospace), x.Id) }));
                        this.Append(this.CreateParagraph(groupsList));
                    }
                    if (group.IncludingGroups.Count > 0)
                    {
                        this.AppendChild(StringToParagraph("This unit is referenced by the following units:"));
                        var groupsList = new ListParagraph(ListParagraphType.Bullet);
                        groupsList.Items.AddRange(group.IncludingGroups.Select(x => new TextParagraph() { new TextRun(x.Name, new(TextRunFormat.Monospace), x.Id) }));
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
                logger.Debug($"Writing {cls.Type} {cls.Name}");

                var (title, anonymousName) = cls.Type switch
                {
                    ClassType.Struct => ("Struct", "Anonymous Struct"),
                    ClassType.Union => ("Union", "Anonymous Union"),
                };

                this.WritePotentiallyAnonymousHeading(title, cls.Name, anonymousName, cls.Id, headingLevel);

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

                        var bookmark = this.bookmarkManager.CreateBookmark(variable.Id);
                        var name = this.ParagraphElementsToRuns(variable.Type).ToList();
                        name.Add(new Run(new Text(" ").PreserveSpace()));
                        name.Add(bookmark.start);
                        name.Add(new Run(new Text(variable.Name ?? "")));
                        name.Add(bookmark.end);
                        if (variable.ArgsString != null || bitfield != null)
                        {
                            name.Add(new Run(new Text($"{variable.ArgsString}{bitfield}").PreserveSpace()));
                        }

                        table.AppendRow(name, this.CreateDescriptions(variable.Descriptions));
                    }
                }
            }
        }

        private void WriteEnums(List<EnumDoc> enums, int headingLevel)
        {
            foreach (var @enum in enums)
            {
                logger.Debug($"Writing Enum {@enum.Name}");

                this.WritePotentiallyAnonymousHeading("Enum", @enum.Name, "Anonymous Enum", @enum.Id, headingLevel);

                this.WriteDescriptions(@enum.Descriptions);

                if (@enum.Values.Count > 0)
                {
                    this.WriteMiniHeading("Values");
                    var table = this.AppendChild(this.CreateTable(StyleManager.ParameterTableStyleId).AddColumns(2));
                    foreach (var value in @enum.Values)
                    {
                        var bookmark = this.bookmarkManager.CreateBookmark(value.Id);
                        var name = new OpenXmlElement[]
                        {
                            bookmark.start,
                            new Run(new Text(value.Name)),
                            bookmark.end,
                        };
                        table.AppendRow(name, this.CreateDescriptions(value.Descriptions));
                    }
                }
            }
        }

        private void WriteTypedefs(List<TypedefDoc> typedefs, int headingLevel)
        {
            foreach (var typedef in typedefs)
            {
                logger.Debug($"Writing typedef {typedef.Name}");

                this.WriteHeading("Typedef", typedef.Name, typedef.Id, headingLevel);

                this.WriteMiniHeading("Definition");
                var paragraph = this.AppendChild(new Paragraph().LeftAlign().ApplyStyle(StyleManager.CodeStyleId));
                paragraph.AppendChild(new Run(new Text("typedef ").PreserveSpace()));
                paragraph.Append(this.ParagraphElementsToRuns(typedef.Type));
                paragraph.AppendChild(new Run(new Text(" " + typedef.Name).PreserveSpace()));

                this.WriteDescriptions(typedef.Descriptions);
            }
        }

        private void WriteGlobalVariables(List<VariableDoc> variables, int headingLevel)
        {
            foreach (var variable in variables)
            {
                logger.Debug($"Writing variable {variable.Name}");

                this.WriteHeading("Global variable", variable.Name!, variable.Id, headingLevel);

                this.WriteMiniHeading("Definition");
                var paragraph = this.AppendChild(new Paragraph().LeftAlign().ApplyStyle(StyleManager.CodeStyleId));
                paragraph.Append(this.ParagraphElementsToRuns(variable.Type));
                paragraph.Append(new Run(new Text($" {variable.Name}").PreserveSpace()));
                if (variable.Initializer.Count > 0)
                {
                    paragraph.AppendChild(new Run(new Text(" ").PreserveSpace()));
                    paragraph.Append(this.ParagraphElementsToRuns(variable.Initializer));
                }

                this.WriteDescriptions(variable.Descriptions);
            }
        }

        private void WriteMacros(List<MacroDoc> macros, int headingLevel)
        {
            foreach (var macro in macros)
            {
                logger.Debug($"Writing macro {macro.Name}");

                this.WriteHeading("Macro", macro.Name, macro.Id, headingLevel);

                this.WriteMiniHeading($"Definition");

                string? paramsStr = macro.HasParameters
                    ? "(" + string.Join(", ", macro.Parameters.Select(x => x.Name)) + ")"
                    : null;
                var paragraph = this.AppendChild(new Paragraph(new Run(new Text($"#define {macro.Name}{paramsStr} ")
                    .PreserveSpace())).ApplyStyle(StyleManager.CodeStyleId).LeftAlign());

                // If it's a multi-line macro declaration, push it onto a new line and add \ to the first line
                if (macro.Initializer.Count > 0 && macro.Initializer[0].Text.StartsWith(" ") &&
                    macro.Initializer.Any(x => x.Text.Contains("\n")))
                {
                    paragraph.AppendChild(new Run(new Text("\\"), new Break()));
                }
                paragraph.Append(this.ParagraphElementsToRuns(macro.Initializer));

                this.WriteDescriptions(macro.Descriptions);

                if (macro.Parameters.Count > 0 && macro.Parameters.Any(x => x.Description.Count > 0))
                {
                    this.WriteMiniHeading("Parameters");

                    var table = this.AppendChild(this.CreateTable(StyleManager.ParameterTableStyleId).AddColumns(2));
                    foreach (var parameter in macro.Parameters)
                    {
                        table.AppendRow(parameter.Name, this.CreateParagraphs(parameter.Description));
                    }
                }

                this.WriteReturnDescription(macro.ReturnDescriptions);
            }
        }

        private void WriteFunctions(List<FunctionDoc> functions, int headingLevel)
        {
            foreach (var function in functions)
            {
                logger.Debug($"Writing function {function.Name}");

                this.WriteHeading("Function", function.Name, function.Id, headingLevel);

                this.WriteMiniHeading("Signature");
                var signatureParagraph = this.AppendChild(new Paragraph().LeftAlign().ApplyStyle(StyleManager.CodeStyleId));
                Run NewRun(OpenXmlElement e) => signatureParagraph.AppendChild(new Run(e));

                signatureParagraph.Append(this.ParagraphElementsToRuns(function.ReturnType));
                NewRun(new Text($" {function.Name}").PreserveSpace());

                if (function.Parameters.Count > 0)
                {
                    NewRun(new Text("("));
                    for (int i = 0; i < function.Parameters.Count;i++)
                    {
                        string comma = i == function.Parameters.Count - 1 ? "" : ",";
                        var run = NewRun(new Break());
                        run.AppendChild(new Text("    ").PreserveSpace());
                        signatureParagraph.Append(this.ParagraphElementsToRuns(function.Parameters[i].Type));
                        string space = signatureParagraph.InnerText.EndsWith(" *") ? "" : " ";
                        NewRun(new Text($"{space}{function.Parameters[i].Name}{comma}").PreserveSpace());
                    }

                    NewRun(new Text(")"));
                }
                else
                {
                    NewRun(new Text(function.ArgsString));
                }

                this.WriteDescriptions(function.Descriptions);

                if (function.Parameters.Count > 0 && function.Parameters.Any(x => x.Description.Count > 0))
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
                        table.AppendRow(parameter.Name, this.CreateParagraphs(parameter.Description), inOut);
                    }
                }

                this.WriteReturnDescription(function.ReturnDescriptions);

                void CreateReferences(List<FunctionDoc> references, string blurb)
                {
                    if (references.Count > 0)
                    {
                        this.AppendChild(new Paragraph(new Run(new Text(blurb))));
                        var list = new ListParagraph(ListParagraphType.Bullet);
                        list.Items.AddRange(references.Select(x => new TextParagraph() { new TextRun(x.Name, new(TextRunFormat.Monospace), x.Id) }));
                        this.Append(this.CreateParagraph(list));
                    }
                }

                if (function.References.Count > 0 || function.ReferencedBy.Count > 0)
                {
                    this.WriteMiniHeading("References");
                    CreateReferences(function.References, "This function references the following functions: ");
                    CreateReferences(function.ReferencedBy, "This function is referenced by the following functions: ");
                }
            }
        }

        private void WriteReturnDescription(ReturnDescriptions returnDescriptions)
        {
            if (returnDescriptions.Description.Any(x => !x.IsEmpty))
            {
                this.WriteMiniHeading("Returns");
                this.Append(this.CreateParagraphs(returnDescriptions.Description));
            }

            if (returnDescriptions.Values.Count > 0)
            {
                this.WriteMiniHeading("Return values");
                var table = this.AppendChild(this.CreateTable(StyleManager.ParameterTableStyleId).AddColumns(2));
                foreach (var returnValue in returnDescriptions.Values)
                {
                    table.AppendRow(returnValue.Name, this.CreateParagraphs(returnValue.Description));
                }
            }
        }

        private void WritePotentiallyAnonymousHeading(string? prefix, string text, string anonymousText, string id, int headingLevel)
        {
            if (text.StartsWith("@"))
            {
                this.WriteHeading(null, anonymousText, id, headingLevel);
            }
            else
            {
                this.WriteHeading(prefix, text, id, headingLevel);
            }
        }

        private void WriteHeading(string? prefix, string text, string id, int headingLevel)
        {
            var paragraph = this.AppendChild(new Paragraph());

            if (prefix != null)
            {
                paragraph.Append(RemoveCaps(new Run(new Text(prefix + " ").PreserveSpace())));
            }

            var run = new Run(new Text(text));
            RemoveCaps(run);
            var (bookmarkStart, bookmarkEnd) = this.bookmarkManager.CreateBookmark(id);
            paragraph.Append(bookmarkStart, run, bookmarkEnd);

            // The style might not have enough spacing, so force this
            paragraph.ApplyStyle($"Heading{headingLevel}")
                .WithProperties(x => x.SpacingBetweenLines = new SpacingBetweenLines() { Before = $"{8 * 20}" });

            // If the style specifies all-caps or small-caps, undo this
            static Run RemoveCaps(Run run) => run.WithProperties(x =>
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
            foreach (var p in this.CreateParagraphs(descriptions.DetailedDescription))
            {
                yield return p;
            }
        }

        private IEnumerable<OpenXmlElement> CreateParagraphs(IEnumerable<IParagraph> paragraphs)
        {
            foreach (var paragraph in paragraphs)
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
                    switch (textParagraph.Type)
                    {
                        case TextParagraphType.Normal:
                            break;
                        case TextParagraphType.Warning:
                            paragraph.ApplyStyle(StyleManager.WarningStyleId);
                            break;
                        case TextParagraphType.Preformatted:
                            paragraph.ApplyStyle(StyleManager.CodeStyleId);
                            break;
                        case TextParagraphType.Note:
                            paragraph.ApplyStyle(StyleManager.NoteStyleId);
                            break;
                        case TextParagraphType.BlockQuote:
                            paragraph.ApplyStyle(StyleManager.BlockQuoteStyleId);
                            break;
                    }
                    if (textParagraph.HasHorizontalRuler)
                    {
                        paragraph.WithProperties(x =>
                        {
                            x.ParagraphBorders ??= new ParagraphBorders();
                            x.ParagraphBorders.BottomBorder = new BottomBorder()
                            {
                                Val = BorderValues.Single,
                                Color = "auto",
                                Space = 1,
                                Size = 6,
                            };
                        });
                    }
                    yield return paragraph;
                    break;

                case TitleParagraph titleParagraph:
                    yield return new Paragraph(this.ParagraphElementsToRuns(titleParagraph)).ApplyStyle(StyleManager.ParHeadingStyleId);
                    break;

                case ListParagraph listParagraph:
                    foreach (var p in this.CreateListParagraph(listParagraph))
                    {
                        yield return p;
                    }
                    break;

                case DefinitionListParagraph definitionListParagraph:
                    foreach (var p in this.CreateDefinitionListParagraph(definitionListParagraph))
                    {
                        yield return p;
                    }
                    break;

                case CodeParagraph codeParagraph:
                    yield return this.CreateCodeParagraph(codeParagraph);
                    break;

                case DotParagraph dotParagraph:
                    foreach (var p in this.CreateDotParagraph(dotParagraph))
                    {
                        yield return p;
                    }
                    break;

                case ImageElement imageElement:
                    foreach (var p in this.CreateImageParagraph(this.imageManager.CreateImage(imageElement.FilePath, imageElement.Dimensions), imageElement.Caption))
                    {
                        yield return p;
                    }
                    break;

                case TableDoc table:
                    foreach (var p in this.CreateTable(table))
                    {
                        yield return p;
                    }
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

            paragraph.Append(this.ParagraphElementsToRuns(textParagraph));
            
            return paragraph;
        }

        private IEnumerable<OpenXmlElement> ParagraphElementsToRuns(IEnumerable<IParagraphElement> paragraphElement)
        {
            foreach (var element in paragraphElement)
            {
                switch (element)
                {
                    case TextRun textRun:
                        foreach (var p in this.ProcessTextRun(textRun))
                        {
                            yield return p;
                        }
                    break;

                    case UrlLink link:
                    {
                        var rel = this.doc.MainDocumentPart!.AddHyperlinkRelationship(new Uri(link.Url), isExternal: true);
                        var hyperlink = new Hyperlink() { Id = rel.Id };
                        hyperlink.Append(this.ParagraphElementsToRuns(link).ApplyRunStyle(StyleManager.HyperlinkStyleId));
                        yield return hyperlink;
                    }
                    break;

                    case ImageElement imageElement:
                    {
                        var image = this.imageManager.CreateImage(imageElement.FilePath, imageElement.Dimensions);
                        if (image != null)
                        {
                            yield return new Run(image);
                        }
                    }
                    break;

                    default:
                        throw new GeneratorException($"Unexpected IParagraphElement type {element} ({element.GetType()})");
                }
            }
        }

        private IEnumerable<OpenXmlElement> ProcessTextRun(TextRun textRun)
        {
            var run = new Run(StringToElements(textRun.Text));

            int fontSize = this.styleManager.DefaultFontSize;
            run.WithProperties(x =>
            {
                x.Bold = textRun.Properties.Format.HasFlag(TextRunFormat.Bold) ? new Bold() : null;
                x.Italic = textRun.Properties.Format.HasFlag(TextRunFormat.Italic) ? new Italic() : null;
                x.Strike = textRun.Properties.Format.HasFlag(TextRunFormat.Strikethrough) ? new Strike() : null;
                x.Underline = textRun.Properties.Format.HasFlag(TextRunFormat.Underline) ? new Underline() { Val = UnderlineValues.Single } : null;
                x.VerticalTextAlignment = textRun.Properties.VerticalPosition switch
                {
                    TextRunVerticalPosition.Baseline => null,
                    TextRunVerticalPosition.Superscript => new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript },
                    TextRunVerticalPosition.Subscript => new VerticalTextAlignment() { Val = VerticalPositionValues.Subscript },
                };
            });
            if (textRun.Properties.Format.HasFlag(TextRunFormat.Monospace))
            {
                fontSize -= 2;
                run.WithProperties(x => x.RunFonts = new RunFonts() { Ascii = "Consolas" });
            }
            if (textRun.Properties.IsSmall)
            {
                fontSize -= 2;
            }

            if (fontSize != this.styleManager.DefaultFontSize)
            {
                run.WithProperties(x => x.FontSize = new FontSize() { Val = fontSize.ToString() });
            }

            OpenXmlElement result = textRun.ReferenceId == null
                ? run
                : this.bookmarkManager.CreateLink(textRun.ReferenceId, run);
            if (textRun.AnchorId == null)
            {
                yield return result;
            }
            else
            {
                var bookmark = this.bookmarkManager.CreateBookmark(textRun.AnchorId);
                yield return bookmark.start;
                yield return result;
                yield return bookmark.end;
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

        private IEnumerable<OpenXmlElement> CreateDefinitionListParagraph(DefinitionListParagraph listParagraph, int level = 0)
        {
            int numberId = this.listManager.CreateList(ListManager.DefinitionListStyleId);

            foreach (var entry in listParagraph.Entries)
            {
                var termParagraph = this.CreateTextParagraph(entry.Term);
                termParagraph.ParagraphProperties = new ParagraphProperties()
                {
                    ParagraphStyleId = new ParagraphStyleId() { Val = StyleManager.DefinitionTermStyleId },
                    NumberingProperties = new NumberingProperties()
                    {
                        NumberingLevelReference = new NumberingLevelReference() { Val = level },
                        NumberingId = new NumberingId() { Val = numberId },
                    },
                };
                yield return termParagraph;

                foreach (var descriptionItem in entry.Description)
                {
                    // This is probably a TextParagraph, or another DefinitionListParagraph

                    switch (descriptionItem)
                    {
                        case TextParagraph textParagraph:
                            var paragraph = this.CreateTextParagraph(textParagraph);
                            paragraph.ParagraphProperties = new ParagraphProperties()
                            {
                                NumberingProperties = new NumberingProperties()
                                {
                                    NumberingLevelReference = new NumberingLevelReference() { Val = level + 1 },
                                    NumberingId = new NumberingId() { Val = numberId },
                                },
                            };
                            yield return paragraph;
                            break;

                        case DefinitionListParagraph definitionList:
                            foreach (var p in this.CreateDefinitionListParagraph(definitionList, level + 1))
                            {
                                yield return p;
                            }
                            break;

                        default:
                            foreach (var p in this.CreateParagraph(descriptionItem))
                            {
                                yield return p;
                            }
                            break;
                    }
                }
            }
        }

        private Paragraph CreateCodeParagraph(CodeParagraph codeParagraph)
        {
            var paragraph = new Paragraph().ApplyStyle(StyleManager.CodeListingStyleId).LeftAlign();

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
        
        private IEnumerable<OpenXmlElement> CreateDotParagraph(DotParagraph dotParagraph)
        {
            byte[]? output = this.dotRenderer.TryRender(dotParagraph.Contents);
            return output != null
                ? this.CreateImageParagraph(this.imageManager.CreateImage(output, ImagePartType.Png, dotParagraph.Dimensions), dotParagraph.Caption)
                : Enumerable.Empty<OpenXmlElement>();
        }

        private IEnumerable<OpenXmlElement> CreateImageParagraph(OpenXmlElement? image, string? caption)
        {
            if (image == null)
            {
                yield break;
            }

            yield return new Paragraph(new Run(image));

            if (!string.IsNullOrEmpty(caption))
            {
                var paragraph = new Paragraph().ApplyStyle(StyleManager.CaptionStyleId);
                paragraph.AppendChild(new Run(new Text("Figure ").PreserveSpace()));
                paragraph.AppendChild(new SimpleField()
                {
                    Instruction = " SEQ Figure \\* ARABIC ",
                    Dirty = true,
                });
                paragraph.AppendChild(new Run(new Text(" ").PreserveSpace()));
                paragraph.AppendChild(new Run(StringToElements(caption)));
                yield return paragraph;
            }
        }

        private IEnumerable<OpenXmlElement> CreateTable(TableDoc tableDoc)
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

            yield return table;

            if (tableDoc.Caption != null)
            {
                var paragraph = this.CreateTextParagraph(tableDoc.Caption);
                paragraph.ApplyStyle(StyleManager.CaptionStyleId);
                var newChildren = new List<OpenXmlElement>()
                {
                    new Run(new Text("Table ").PreserveSpace()),
                    new SimpleField()
                    {
                        Instruction = " SEQ TABLE \\* ARABIC ",
                        Dirty = true,
                    },
                    new Run(new Text(" ").PreserveSpace()),
                };
                if (tableDoc.Id != null)
                {
                    var bookmark = this.bookmarkManager.CreateBookmark(tableDoc.Id);
                    newChildren.Add(bookmark.start);
                    paragraph.AppendChild(bookmark.end);
                }
                for (int i = 0; i < newChildren.Count; i++)
                {
                    paragraph.InsertAt(newChildren[i], i + 1); // After the ParagraphProperties
                }

                yield return paragraph;
            }
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
