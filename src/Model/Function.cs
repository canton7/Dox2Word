using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class Function
    {
        public string Name { get; set; } = null!;
        public Paragraph BriefDescription { get; set; } = null!;
        public List<Paragraph> DetailedDescription { get; } = new();
        public string ReturnType { get; set; } = null!;
        public string ArgsString { get; set; } = null!;
        public List<FunctionParameter> FunctionParameters { get; } = new();
    }
}