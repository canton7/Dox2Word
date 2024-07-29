using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class EnumDoc
    {
        public string Id { get; set; } = null!;
        public string Name { get; set; } = null!;
        public List<EnumValueDoc> Values { get; } = [];
        public Descriptions Descriptions { get; set; } = null!;
    }
}