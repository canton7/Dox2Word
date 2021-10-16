using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dox2Word.Model
{
    public class Group
    {
        public string Name { get; init; } = null!;

        public List<Group> SubGroups { get; } = new();

        public List<Function> Functions { get; } = new();
    }
}
