namespace Dox2Word.Model
{
    public class DotParagraph : IParagraph
    {
        public string Contents { get; private set; }

        public string? Caption { get; set; }

        public bool IsEmpty => string.IsNullOrEmpty(this.Contents);

        public DotParagraph(string contents) => this.Contents = contents;

        public void TrimTrailingWhitespace() => this.Contents = this.Contents.TrimEnd(' ');
    }
}
