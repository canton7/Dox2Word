namespace Dox2Word.Model
{
    public class ImageParagraph : IParagraph
    {
        public string FilePath { get; }
        public string? Caption { get; }
        public bool IsInline { get; }

        public bool IsEmpty => false;

        public ImageParagraph(string filePath, string? caption, bool isInline)
        {
            this.FilePath = filePath;
            this.Caption = caption;
            this.IsInline = isInline;
        }

        public void TrimTrailingWhitespace() { }
    }
}
