using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class Group
    {
        public string Id { get; set; } = null!;
        public string Name { get; set; } = null!;
        public Descriptions Descriptions { get; set; } = null!;
        public List<Group> SubGroups { get; } = [];
        public List<Group> IncludedGroups { get; } = [];
        public List<Group> IncludingGroups { get; } = [];
        public List<FileDoc> Files { get; } = [];
        public List<ClassDoc> Classes { get; } = [];
        public List<EnumDoc> Enums { get; } = [];
        public List<TypedefDoc> Typedefs { get; } = [];
        public List<VariableDoc> GlobalVariables { get; } = [];
        public List<MacroDoc> Macros { get; } = [];
        public List<FunctionDoc> Functions { get; } = [];
    }
}
