using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Dox2Word.Model
{
    public interface ITextRun { }

    [Flags]
    public enum TextRunFormat
    {
        None,
        Bold = 1,
        Italic = 2,
        Monospace = 4,
    }

    public class TextRun : ITextRun
    {
        public TextRunFormat Format { get; }

        public string Text { get; }

        public TextRun(string text, TextRunFormat format)
        {
            this.Text = text;
            this.Format = format;
        }
    }

    public enum ListTextRunType
    {
        Number,
        Bullet,
    }

    public class ListTextRun : ITextRun
    {
        public ListTextRunType Type { get; }

        public List<ListTextRunItem> Items { get; } = new();

        public ListTextRun(ListTextRunType type)
        {
            this.Type = type;
        }
    }

    public class ListTextRunItem : TextParagraph
    {
    }
}