namespace Dox2Word.Model
{
    public class VariableDoc
    {
        public string Name { get; set; } = null!;
        public string Type { get; set; } = null!;
        public string Definition { get; set; } = null!;
        public Descriptions Descriptions { get; set; } = null!;
        public string? Initializer { get; internal set; }
        public string? Bitfield { get; set; }
    }
}
