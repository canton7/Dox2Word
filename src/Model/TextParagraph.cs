using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dox2Word.Model
{
    public interface IParagraph
    {
        int Count { get; }
    }

    public enum ParagraphType
    {
        Normal,
        Warning,
    }

    public class TextParagraph : Collection<TextRun>, IParagraph
    {
        public ParagraphType Type { get; }

        public TextParagraph(ParagraphType type = ParagraphType.Normal)
        {
            this.Type = type;
        }
    }

    public enum ListParagraphType
    {
        Number,
        Bullet,
    }

    public class ListParagraph : IParagraph
    {
        public ListParagraphType Type { get; }

        public List<IParagraph> Items { get; } = new();

        int IParagraph.Count => this.Items.Count;

        public ListParagraph(ListParagraphType type)
        {
            this.Type = type;
        }
    }
}
