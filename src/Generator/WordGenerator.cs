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
            this.WriteGlobalVariables(group.GlobalVariables, headingLevel + 1);
            this.WriteFunctions(group.Functions, headingLevel + 1);

            foreach (var subGroup in group.SubGroups)
            {
                this.WriteGroup(subGroup, headingLevel + 1);
            }
        }

        private void WriteClasses(List<Class> classes, int headingLevel)
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
                    var table = CreateTable();
                    this.bodyElements.Add(table);
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

        private void WriteGlobalVariables(List<Variable> variables, int headingLevel)
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

        private void WriteFunctions(List<Function> functions, int headingLevel)
        {
            if (functions.Count == 0)
                return;

            this.WriteHeading("Functions", headingLevel);

            foreach (var function in functions)
            {
                this.WriteHeading(function.Name, headingLevel + 1);

                var paragraph = this.AppendChild(new Paragraph());
                var run = paragraph.AppendChild(new Run(new Text(function.Definition), new Text(function.ArgsString)));
                run.FormatCode();

                this.Append(this.CreateDescriptions(function.Descriptions));

                if (function.Parameters.Count > 0)
                {
                    var table = CreateTable();
                    this.bodyElements.Add(table);
                    foreach (var parameter in function.Parameters)
                    {
                        var row = table.AppendChild(new TableRow());
                        var nameCell = row.AppendChild(CreateTableCell());
                        nameCell.AppendChild(new Paragraph(new Run(new Text(parameter.Name)).FormatCode()));
                        var descriptionCell = row.AppendChild(CreateTableCell());
                        descriptionCell.Append(this.CreateParagraph(parameter.Description));
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
                }
            }
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

        private static Table CreateTable()
        {
            var table = new Table();

            var tableProperties = table.AppendChild(new TableProperties());
            var tableBorders = tableProperties.AppendChild(new TableBorders());

            var topBorder = tableBorders.AppendChild(new TopBorder());
            topBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);

            var rightBorder = tableBorders.AppendChild(new RightBorder());
            rightBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);

            var bottomBorder = tableBorders.AppendChild(new BottomBorder());
            bottomBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);

            var leftBorder = tableBorders.AppendChild(new LeftBorder());
            leftBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);

            var insideHorizontalBorder = tableBorders.AppendChild(new InsideHorizontalBorder());
            insideHorizontalBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);

            var insideVerticalBorder = tableBorders.AppendChild(new InsideVerticalBorder());
            insideVerticalBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);

            return table;
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
