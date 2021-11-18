using System;

namespace Dox2Word.Model
{
    [Flags]
    public enum TextRunFormat
    {
        None,
        Bold = 1,
        Italic = 2,
        Monospace = 4,
    }

    public class TextRun
    {
        public TextRunFormat Format { get; }

        public string Text { get; set; }

        public string? ReferenceId { get; }

        public TextRun(string text, TextRunFormat format = TextRunFormat.None, string? referenceId = null)
        {
            this.Text = text;
            this.Format = format;
            this.ReferenceId = referenceId;
        }
    }
}