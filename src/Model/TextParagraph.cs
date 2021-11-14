using System.Collections.ObjectModel;
using System.Linq;

namespace Dox2Word.Model
{
    public enum TextParagraphType
    {
        Normal,
        Warning,
    }

    public enum TextParagraphAlignment
    {
        Default,
        Left,
        Center,
        Right,
    }

    public class TextParagraph : Collection<TextRun>, IParagraph
    {
        public TextParagraphType Type { get; }

        public TextParagraphAlignment Alignment { get; }

        public bool IsEmpty => this.Count == 0;

        public TextParagraph(TextParagraphType type = TextParagraphType.Normal, TextParagraphAlignment alignment = TextParagraphAlignment.Default)
        {
            this.Type = type;
            this.Alignment = alignment;
        }

        public void TrimTrailingWhitespace()
        {
            if (this.LastOrDefault() is { } run && run.Text.EndsWith(" "))
            {
                run.Text = run.Text.TrimEnd(' ');
            }
        }
    }
}
