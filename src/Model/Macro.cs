using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dox2Word.Model
{
    public class Macro
    {
        public string Name { get; set; } = null!;
        public Descriptions Descriptions { get; set; } = null!;
        public TextParagraph ReturnDescription { get; set; } = null!;
        public string Initializer { get; set; } = null!;
        public List<Parameter> Parameters { get; } = new();
    }
}
