namespace Dox2Word.Model
{
    public class Parameter
    {
        public string Name { get; set; } = null!;
        public string? Type { get; set; }
        public TextParagraph Description { get; set; } = null!;
    }
}