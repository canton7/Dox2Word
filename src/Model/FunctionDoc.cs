using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class FunctionDoc
    {
        public string Id { get; set; } = null!;
        public string Name { get; set; } = null!;
        public Descriptions Descriptions { get; set; } = null!;
        public List<TextRun> ReturnType { get; set; } = null!;
        public ReturnDescriptions ReturnDescriptions { get; set; } = null!;
        public string Definition { get; set; } = null!;
        public string ArgsString { get; set; } = null!;
        public List<ParameterDoc> Parameters { get; } = new();
    }
}