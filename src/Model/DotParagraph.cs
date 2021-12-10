namespace Dox2Word.Model
{
    public class DotParagraph : IParagraph
    {
        public string Contents { get; private set; }

        public string? Caption { get; }
        public ImageDimensions Dimensions { get; }

        public bool IsEmpty => string.IsNullOrEmpty(this.Contents);

        public DotParagraph(string contents, string? caption, ImageDimensions dimensions)
        {
            this.Contents = contents;
            this.Caption = caption;
            this.Dimensions = dimensions;
        }

        public void TrimTrailingWhitespace() => this.Contents = this.Contents.TrimEnd(' ');
    }
}
