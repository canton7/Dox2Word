using System.Collections.Generic;
using System.Collections.ObjectModel;

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

    public enum ParagraphAlignment
    {
        Default,
        Left,
        Center,
        Right,
    }

    public class TextParagraph : Collection<TextRun>, IParagraph
    {
        public ParagraphType Type { get; }

        public ParagraphAlignment Alignment { get; }

        public TextParagraph(ParagraphType type = ParagraphType.Normal, ParagraphAlignment alignment = ParagraphAlignment.Default)
        {
            this.Type = type;
            this.Alignment = alignment;
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

    public class CodeParagraph : IParagraph
    {
        public List<string> Lines { get; } = new();

        int IParagraph.Count => this.Lines.Count;
    }
}
