namespace Dox2Word.Model
{
    public class FunctionParameter
    {
        public string Name { get; set; } = null!;
        public string Type { get; set; } = null!;
        public Paragraph Description { get; set; } = null!;
    }
}