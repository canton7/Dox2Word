namespace Dox2Word.Model
{
    public enum ImageDimensionUnit
    {
        Px,
        Cm,
        Percent,
        Inch,
    }

    public struct ImageDimension
    {
        public double Value { get; }
        public ImageDimensionUnit Unit { get; }

        public ImageDimension(double value, ImageDimensionUnit unit)
        {
            this.Value = value;
            this.Unit = unit;
        }
    }

    public struct ImageDimensions
    {
        public ImageDimension? Width { get; }
        public ImageDimension? Height { get; }

        public ImageDimensions(ImageDimension? width, ImageDimension? height)
        {
            this.Width = width;
            this.Height = height;
        }
    }

    public class ImageElement : IParagraph, IParagraphElement
    {
        public string FilePath { get; }
        public string? Caption { get; }
        public ImageDimensions Dimensions { get; }

        public bool IsEmpty => false;

        public ImageElement(string filePath, string? caption, ImageDimensions dimensions)
        {
            this.FilePath = filePath;
            this.Caption = caption;
            this.Dimensions = dimensions;
        }

        public void TrimTrailingWhitespace() { }
    }
}
