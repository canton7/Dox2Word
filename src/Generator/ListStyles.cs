using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Dox2Word.Model;

namespace Dox2Word.Generator
{
    public class ListStyles
    {
        private readonly NumberingDefinitionsPart numberingPart;

        private readonly int bulletAbstractNumId;
        private readonly int numberedAbstractNumId;

        public ListStyles(NumberingDefinitionsPart numberingPart)
        {
            this.numberingPart = numberingPart;

            this.bulletAbstractNumId = this.FindOrCreateBulletStyle();
            this.numberedAbstractNumId = this.FindOrCreateNumberedStyle();
        }

        private int FindOrCreateBulletStyle()
        {
            int? abstractNumberId = this.numberingPart.Numbering.Elements<AbstractNum>()
                .FirstOrDefault(x => x.MultiLevelType?.Val?.Value == MultiLevelValues.HybridMultilevel &&
                    x.Elements<Level>().FirstOrDefault()?.NumberingFormat?.Val?.Value == NumberFormatValues.Bullet)
                ?.AbstractNumberId?.Value;

            return abstractNumberId != null
                ? abstractNumberId.Value
                : this.CreateAbstractNum("BulletStyle.xml");
        }

        private int FindOrCreateNumberedStyle()
        {
            int? abstractNumberId = this.numberingPart.Numbering.Elements<AbstractNum>()
                .FirstOrDefault(x => x.MultiLevelType?.Val?.Value == MultiLevelValues.HybridMultilevel &&
                    x.Elements<Level>().FirstOrDefault()?.NumberingFormat?.Val?.Value == NumberFormatValues.Decimal)
                ?.AbstractNumberId?.Value;

            return abstractNumberId != null
                ? abstractNumberId.Value
                : this.CreateAbstractNum("NumberedStyle.xml");
        }

        private int CreateAbstractNum(string fileName)
        {
            int abstractNumId = this.numberingPart.Numbering.Elements<AbstractNum>().MaxOrDefault(x => x.AbstractNumberId?.Value ?? 0, 0) + 1;

            var abstractNum = new AbstractNum() { AbstractNumberId = abstractNumId };
            using var sr = new StreamReader(typeof(ListStyles).Assembly.GetManifestResourceStream($"Dox2Word.Generator.{fileName}"));
            abstractNum.InnerXml = sr.ReadToEnd();

            // Insert an AbstractNum into the numbering part numbering list.  The order seems to matter or it will not pass the 
            // Open XML SDK Productity Tools validation test.  AbstractNum comes first and then NumberingInstance and we want to
            // insert this AFTER the last AbstractNum and BEFORE the first NumberingInstance or we will get a validation error.
            this.numberingPart.Numbering.InsertAfter(abstractNum, this.numberingPart.Numbering.Elements<AbstractNum>().LastOrDefault());

            return abstractNumId;
        }

        public int CreateList(ListParagraphType type)
        {
            int abstractNumId = type switch
            {
                ListParagraphType.Bullet => this.bulletAbstractNumId,
                ListParagraphType.Number => this.numberedAbstractNumId,
            };

            int numberId = this.numberingPart.Numbering.Elements<NumberingInstance>().MaxOrDefault(x => x.NumberID?.Value ?? 0, 0) + 1;

            // Insert an NumberingInstance into the numbering part numbering list.  The order seems to matter or it will not pass the 
            // Open XML SDK Productity Tools validation test.  AbstractNum comes first and then NumberingInstance and we want to
            // insert this AFTER the last NumberingInstance and AFTER all the AbstractNum entries or we will get a validation error.
            var numberingInstance = new NumberingInstance() { NumberID = numberId };
            numberingInstance.AppendChild(new AbstractNumId() { Val = abstractNumId });

            this.numberingPart.Numbering.InsertAfter(numberingInstance, this.numberingPart.Numbering.Elements<NumberingInstance>().LastOrDefault());

            return numberId;
        }
    }
}
