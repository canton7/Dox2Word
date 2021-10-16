namespace Dox2Word.Model
{
    public class TextRun
    {
        public string Text { get; set; } = null!;
        public string? Format { get; set; }

        public TextRun(string text, string? format = null)
        {
            this.Text = text;
            this.Format = format;
        }
    }
}