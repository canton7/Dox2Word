using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dox2Word.Model
{
    public class Descriptions
    {
        public Paragraph BriefDescription { get; set; } = null!;
        public List<Paragraph> DetailedDescription { get; } = new();
    }
}
