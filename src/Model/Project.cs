using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class Project
    {
        public Dictionary<string, ProjectOption> Options { get; } = [];

        public Dictionary<string, Group> AllGroups { get; set; } = null!;
        public Dictionary<string, FunctionDoc> AllFunctions { get; } = [];

        public List<Group> RootGroups { get; } = [];
    }
}
