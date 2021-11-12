using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Dox2Word.Model
{
    public interface IParagraph
    {
        int Count { get; }

        void Trim();
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

        public void Trim()
        {
            if (this.LastOrDefault() is { } run && run.Text.EndsWith(" "))
            {
                run.Text = run.Text.TrimEnd(' ');
            }
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

        public void Trim() { }
    }

    public class CodeParagraph : IParagraph
    {
        public List<string> Lines { get; } = new();

        int IParagraph.Count => this.Lines.Count;

        public void Trim() { }
    }
}
