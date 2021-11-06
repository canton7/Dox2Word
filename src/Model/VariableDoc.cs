using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dox2Word.Parser.Models;

namespace Dox2Word.Model
{
    public class VariableDoc
    {
        public string Name { get; set; } = null!;
        public string Type { get; set; } = null!;
        public string Definition { get; set; } = null!;
        public Descriptions Descriptions { get; set; } = null!;
        public string? Initializer { get; internal set; }
        public string? Bitfield { get; set; }
    }
}
