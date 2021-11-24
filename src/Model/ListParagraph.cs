using System.Collections.Generic;

namespace Dox2Word.Model
{
    public enum ListParagraphType
    {
        Number,
        Bullet,
    }

    public class ListParagraph : IParagraph
    {
        public ListParagraphType Type { get; }

        public List<IParagraph> Items { get; } = new();

        public bool IsEmpty => this.Items.Count == 0;

        public ListParagraph(ListParagraphType type)
        {
            this.Type = type;
        }

        public void TrimTrailingWhitespace() { }
    }

}
