using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Dox2Word.Model;

namespace Dox2Word.Generator
{
    public class ListManager
    {
        // We use list styles, which are a bit complex
        // 1. Our list numbering scheme is defined as an abstractNum. This has a link to the style name using styleLink
        // 2. We define a num, which gives a numId to that abstractNum
        // 3. We define a style, which references the num above
        // Whenever we insert a new list instance:
        // 1. We define a new abstractNum, which links to the style using numStyleLink
        // 2. We define a new num which gives a numId to that abstractNum

        private readonly NumberingDefinitionsPart numberingPart;
        private readonly StyleManager styleManager;

        public const string BulletStyleId = "DoxBulletList";
        public const string NumberedStyleId = "DoxNumberedList";
        public const string DefinitionListStyleId = "DoxDefinitionList";

        public ListManager(NumberingDefinitionsPart numberingPart, StyleManager styleManager)
        {
            this.numberingPart = numberingPart;
            this.styleManager = styleManager;
        }

        public void EnsureNumbers()
        {
            this.CreateNumberingStyle(BulletStyleId, "Dox Bullet List", "BulletStyle.xml");
            this.CreateNumberingStyle(NumberedStyleId, "Dox Numbered List", "NumberedStyle.xml");
            this.CreateNumberingStyle(DefinitionListStyleId, "Dox Definition List", "DefinitionListStyle.xml");
        }

        private void CreateNumberingStyle(string styleId, string styleName, string filename)
        {
            if (this.styleManager.HasStyleWithId(styleId))
                return;

            var abstractNum = this.CreateAbstractNumFromFile(filename);
            abstractNum.StyleLink = new StyleLink() { Val = styleId };
            int numId = this.CreateNumbering(abstractNum.AbstractNumberId!.Value);

            this.styleManager.AddListStyle(styleId, styleName, numId);
        }

        public int CreateList(ListParagraphType type)
        {
            string styleId = type switch
            {
                ListParagraphType.Bullet => BulletStyleId,
                ListParagraphType.Number => NumberedStyleId,
            };

            return this.CreateList(styleId);
        }

        public int CreateList(string styleId)
        {
            var abstractNum = this.CreateAbstractNum();
            abstractNum.MultiLevelType = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            abstractNum.NumberingStyleLink = new NumberingStyleLink() { Val = styleId };

            int numId = this.CreateNumbering(abstractNum.AbstractNumberId!.Value);
            return numId;
        }

        private AbstractNum CreateAbstractNumFromFile(string filename)
        {
            var abstractNum = this.CreateAbstractNum();
            using var sr = new StreamReader(typeof(ListManager).Assembly.GetManifestResourceStream($"Dox2Word.Generator.{filename}"));
            abstractNum.InnerXml = sr.ReadToEnd();
            
            return abstractNum;
        }

        private AbstractNum CreateAbstractNum()
        {
            int abstractNumId = this.numberingPart.Numbering.Elements<AbstractNum>().MaxOrDefault(x => x.AbstractNumberId?.Value ?? 0, 0) + 1;
            var abstractNum = new AbstractNum() { AbstractNumberId = abstractNumId };

            // Insert an AbstractNum into the numbering part numbering list.  The order seems to matter or it will not pass the 
            // Open XML SDK Productity Tools validation test.  AbstractNum comes first and then NumberingInstance and we want to
            // insert this AFTER the last AbstractNum and BEFORE the first NumberingInstance or we will get a validation error.
            this.numberingPart.Numbering.InsertAfter(abstractNum, this.numberingPart.Numbering.Elements<AbstractNum>().LastOrDefault());

            return abstractNum;
        }

        private int CreateNumbering(int abstractNumId)
        {
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
