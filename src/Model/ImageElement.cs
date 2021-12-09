namespace Dox2Word.Model
{
    public class ImageElement : IParagraph, IParagraphElement
    {
        public string FilePath { get; }
        public string? Caption { get; }

        public bool IsEmpty => false;

        public ImageElement(string filePath, string? caption)
        {
            this.FilePath = filePath;
            this.Caption = caption;
        }

        public void TrimTrailingWhitespace() { }
    }
}
