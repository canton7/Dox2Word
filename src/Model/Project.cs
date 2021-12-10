using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class Project
    {
        public Dictionary<string, ProjectOption> Options { get; } = new();

        public Dictionary<string, Group> AllGroups { get; set; } = null!;
        public Dictionary<string, FunctionDoc> AllFunctions { get; } = new();

        public List<Group> RootGroups { get; } = new();
    }
}
