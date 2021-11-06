using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dox2Word.Model
{
    public class MacroDoc
    {
        public string Name { get; set; } = null!;
        public Descriptions Descriptions { get; set; } = null!;
        public ReturnDescriptions ReturnDescriptions { get; set; } = null!;
        public string Initializer { get; set; } = null!;
        public List<ParameterDoc> Parameters { get; } = new();
    }
}
