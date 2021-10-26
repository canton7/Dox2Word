using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class FunctionDoc
    {
        public string Name { get; set; } = null!;
        public Descriptions Descriptions { get; set; } = null!;
        public string ReturnType { get; set; } = null!;
        public IParagraph ReturnDescription { get; set; } = null!;
        public List<ReturnValueDoc> ReturnValues { get; } = new();
        public string Definition { get; set; } = null!;
        public string ArgsString { get; set; } = null!;
        public List<ParameterDoc> Parameters { get; } = new();
    }
}