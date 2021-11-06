using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class Descriptions
    {
        public IParagraph BriefDescription { get; set; } = null!;
        public List<IParagraph> DetailedDescription { get; } = new();
    }
}
