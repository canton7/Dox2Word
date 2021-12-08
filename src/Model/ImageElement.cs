namespace Dox2Word.Model
{
    public class ImageElement : IParagraphElement
    {
        public string FilePath { get; }
        public string? Caption { get; }

        public ImageElement(string filePath, string? caption)
        {
            this.FilePath = filePath;
            this.Caption = caption;
        }
    }
}
