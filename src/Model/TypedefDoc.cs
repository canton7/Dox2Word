namespace Dox2Word.Model
{
    public class TypedefDoc
    {
        public string Id { get; set; } = null!;
        public string Name { get; set; } = null!;
        public string Type { get; set; } = null!;
        public string Definition { get; set; } = null!;
        public Descriptions Descriptions { get; set; } = null!;
    }
}
