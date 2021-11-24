using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class EnumValueDoc
    {
        public string Id { get; set; } = null!;
        public string Name { get; set; } = null!;
        public List<TextRun> Initializer { get; set; } = null!;
        public Descriptions Descriptions { get; set; } = null!;
    }
}