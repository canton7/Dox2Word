using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class Group
    {
        public string Name { get; set; } = null!;
        public Descriptions Descriptions { get; set; } = null!;
        public List<Group> SubGroups { get; } = new();
        public List<string> Files { get; } = new();
        public List<ClassDoc> Classes { get; } = new();
        public List<EnumDoc> Enums { get; } = new();
        public List<TypedefDoc> Typedefs { get; } = new();
        public List<VariableDoc> GlobalVariables { get; } = new();
        public List<MacroDoc> Macros { get; } = new();
        public List<FunctionDoc> Functions { get; } = new();
    }
}
