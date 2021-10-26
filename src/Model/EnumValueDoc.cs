namespace Dox2Word.Model
{
    public class EnumValueDoc
    {
        public string Name { get; set; } = null!;
        public string? Initializer { get; set; }
        public Descriptions Descriptions { get; set; } = null!;
    }
}