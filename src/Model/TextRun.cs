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
        Strikethrough = 8,
        Underline = 16,
    }

    public enum TextRunVerticalPosition
    {
        Baseline,
        Subscript,
        Superscript,
    }

    public class TextRun : IParagraphElement
    {
        public TextRunProperties Properties { get; }

        public string Text { get; set; }

        public string? ReferenceId { get; }

        public TextRun(string text, TextRunProperties properties = default, string? referenceId = null)
        {
            this.Text = text;
            this.Properties = properties;
            this.ReferenceId = referenceId;
        }
    }

    public struct TextRunProperties
    {
        public TextRunFormat Format { get; }
        public TextRunVerticalPosition VerticalPosition { get; }
        public bool IsSmall { get; }

        public static TextRunProperties Small => new(TextRunFormat.None, TextRunVerticalPosition.Baseline, isSmall: true);

        private TextRunProperties(TextRunFormat format, TextRunVerticalPosition verticalPosition, bool isSmall)
        {
            this.Format = format;
            this.VerticalPosition = verticalPosition;
            this.IsSmall = isSmall;
        }

        public TextRunProperties(TextRunFormat format, TextRunVerticalPosition verticalPosition)
            : this(format, verticalPosition, false)
        {
        }

        public TextRunProperties(TextRunFormat format)
            : this(format, TextRunVerticalPosition.Baseline)
        {
        }

        public TextRunProperties(TextRunVerticalPosition verticalPosition)
           : this(TextRunFormat.None, verticalPosition)
        {
        }

        public TextRunProperties Combine(TextRunProperties other) => new TextRunProperties(
            this.Format | other.Format,
            other.VerticalPosition == TextRunVerticalPosition.Baseline ? this.VerticalPosition : other.VerticalPosition,
            this.IsSmall || other.IsSmall);

        public TextRunProperties Combine(TextRunFormat format) => new TextRunProperties(this.Format | format, this.VerticalPosition, this.IsSmall);
    }
}