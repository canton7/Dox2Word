using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dox2Word.Model
{
    public class Group
    {
        public string Name { get; set; } = null!;
        public Descriptions Descriptions { get; set; } = null!;
        public List<Group> SubGroups { get; } = new();
        public List<string> Files { get; } = new();
        public List<Class> Classes { get; } = new();
        public List<Typedef> Typedefs { get; } = new();
        public List<Variable> GlobalVariables { get; } = new();
        public List<Macro> Macros { get; } = new();
        public List<Function> Functions { get; } = new();
    }
}
