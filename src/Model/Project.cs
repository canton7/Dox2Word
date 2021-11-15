using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class Project
    {
        public List<ProjectOption> Options { get; } = new();

        public List<Group> Groups { get; } = new();
    }
}
