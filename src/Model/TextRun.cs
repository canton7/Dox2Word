using System.Collections.Generic;
using System.Linq;

namespace Dox2Word.Model
{
    public interface ITextRun { }

    public class LiteralTextRun : ITextRun
    {
        public string Text { get; }

        public LiteralTextRun(string text)
        {
            this.Text = text;
        }
    }

    public enum TextRunFormat
    {
        Bold,
        Italic,
        Monospace,
    }

    public class TextRun : ITextRun
    {
        public TextRunFormat Format { get; }

        public List<ITextRun> Children { get; }

        public TextRun(TextRunFormat format, IEnumerable<ITextRun> children)
        {
            this.Format = format;
            this.Children = children.ToList();
        }
    }
}