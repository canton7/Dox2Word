using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class Project
    {
        public List<ProjectOption> Options { get; } = new();

        public Dictionary<string, Group> AllGroups { get; set; } = null!;

        public List<Group> RootGroups { get; } = new();
    }
}
