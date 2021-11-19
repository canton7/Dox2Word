using System.Collections.Generic;

namespace Dox2Word.Model
{
    public enum ClassType
    {
        Struct,
        Union,
    }

    public class ClassDoc
    {
        public ClassType Type { get; set; }
        public string Id { get; set; } = null!;
        public string Name { get; set; } = null!;
        public Descriptions Descriptions { get; set; } = null!;
        public List<VariableDoc> Variables { get; } = new();
    }
}
