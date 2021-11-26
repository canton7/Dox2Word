using System.Collections.Generic;
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

    public interface IParagraphElement
    {
    }

    public class TextParagraph : Collection<IParagraphElement>, IParagraph
    {
        public TextParagraphType Type { get; }

        public TextParagraphAlignment Alignment { get; }

        public bool IsEmpty => this.Count == 0;

        public TextParagraph(TextParagraphType type = TextParagraphType.Normal, TextParagraphAlignment alignment = TextParagraphAlignment.Default)
        {
            this.Type = type;
            this.Alignment = alignment;
        }

        public TextParagraph(TextParagraphAlignment alignment)
            : this(TextParagraphType.Normal, alignment)
        {
        }

        public void AddRange(IEnumerable<TextRun> elements)
        {
            foreach (var element in elements)
            {
                this.Add(element);
            }
        }

        public void TrimTrailingWhitespace()
        {
            if (this.LastOrDefault() is TextRun run && run.Text.EndsWith(" "))
            {
                run.Text = run.Text.TrimEnd(' ');
                if (run.Text.Length == 0)
                {
                    this.Remove(run);
                }
            }
        }
    }
}
