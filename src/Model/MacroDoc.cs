using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class MacroDoc
    {
        public string Name { get; set; } = null!;
        public Descriptions Descriptions { get; set; } = null!;
        public ReturnDescriptions ReturnDescriptions { get; set; } = null!;
        public string Initializer { get; set; } = null!;
        public bool HasParameters { get; set; }
        public List<ParameterDoc> Parameters { get; } = new();
    }
}
