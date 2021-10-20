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

        private readonly List<OpenXmlElement> bodyElements = new();
        private NumberingDefinitionsPart? numberingPart;

        public WordGenerator(Stream stream)
        {
            this.stream = stream;
        }

        public void Generate(Project project)
        {
            this.bodyElements.Clear();

            using var doc = WordprocessingDocument.Open(this.stream, isEditable: true);
            var body = doc.MainDocumentPart!.Document.Body!;
            this.numberingPart = doc.MainDocumentPart.NumberingDefinitionsPart;
            if (this.numberingPart == null)
            {
                this.numberingPart = doc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>("NumberingDefinitionsPart001");
                var element = new Numbering();
                element.Save(this.numberingPart);
            }

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

            this.Append(CreateDescriptions(group.Descriptions));

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

                this.Append(CreateDescriptions(cls.Descriptions));

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
                        descriptionCell.Append(CreateDescriptions(variable.Descriptions));
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

                this.Append(CreateDescriptions(variable.Descriptions));
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

                this.Append(CreateDescriptions(function.Descriptions));

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
                        descriptionCell.AppendChild(CreateTextParagraph(parameter.Description));
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

        private static IEnumerable<OpenXmlElement> CreateDescriptions(Descriptions descriptions)
        {
            yield return CreateTextParagraph(descriptions.BriefDescription);
            foreach (var paragraph in descriptions.DetailedDescription)
            {
                yield return CreateTextParagraph(paragraph);
            }
        }

        private static Paragraph CreateTextParagraph(TextParagraph textParagraph)
        {
            var paragraph = new Paragraph();

            // TODO: Type (warning, etc)

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

            return paragraph;
        }

        private void CreateList(ListTextRun textRun)
        {
            // From https://stackoverflow.com/a/38881677/1086121

            // Insert an AbstractNum into the numbering part numbering list.  The order seems to matter or it will not pass the 
            // Open XML SDK Productity Tools validation test.  AbstractNum comes first and then NumberingInstance and we want to
            // insert this AFTER the last AbstractNum and BEFORE the first NumberingInstance or we will get a validation error.
            int abstractNumberId = this.numberingPart!.Numbering.Elements<AbstractNum>().Count() + 1;
            var abstractLevel = new Level(new NumberingFormat() { Val = NumberFormatValues.Bullet }, new LevelText() { Val = "·" }) { LevelIndex = 0 };
            var abstractNum = new AbstractNum(abstractLevel) { AbstractNumberId = abstractNumberId };

            if (abstractNumberId == 1)
            {
                this.numberingPart.Numbering.Append(abstractNum);
            }
            else
            {
                var lastAbstractNum = this.numberingPart.Numbering.Elements<AbstractNum>().Last();
                this.numberingPart.Numbering.InsertAfter(abstractNum, lastAbstractNum);
            }

            // Insert an NumberingInstance into the numbering part numbering list.  The order seems to matter or it will not pass the 
            // Open XML SDK Productity Tools validation test.  AbstractNum comes first and then NumberingInstance and we want to
            // insert this AFTER the last NumberingInstance and AFTER all the AbstractNum entries or we will get a validation error.
            int numberId = this.numberingPart!.Numbering.Elements<NumberingInstance>().Count() + 1;
            var numberingInstance = new NumberingInstance() { NumberID = numberId };
            numberingInstance.AppendChild(new AbstractNumId() { Val = abstractNumberId });

            if (numberId == 1)
            {
                this.numberingPart.Numbering.Append(numberingInstance);
            }
            else
            {
                var lastNumberingInstance = this.numberingPart.Numbering.Elements<NumberingInstance>().Last();
                this.numberingPart.Numbering.InsertAfter(numberingInstance, lastNumberingInstance);
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
