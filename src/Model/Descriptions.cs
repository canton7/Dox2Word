using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dox2Word.Model
{
    public class Descriptions
    {
        public IParagraph BriefDescription { get; set; } = null!;
        public List<IParagraph> DetailedDescription { get; } = new();
    }
}
