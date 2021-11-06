using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dox2Word.Model
{
    public class ReturnDescriptions
    {
        public IParagraph Description { get; set; } = null!;
        public List<ReturnValueDoc> Values { get; } = new();
    }
}
