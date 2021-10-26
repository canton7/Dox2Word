namespace Dox2Word.Model
{
    public class ParameterDoc
    {
        public string Name { get; set; } = null!;
        public string? Type { get; set; }
        public IParagraph Description { get; set; } = null!;
    }
}