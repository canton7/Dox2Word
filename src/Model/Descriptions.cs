using System.Collections.Generic;
using System.Linq;

namespace Dox2Word.Model
{
    public class Descriptions
    {
        public bool HasDescriptions => !this.BriefDescription.IsEmpty || this.DetailedDescription.Any(x => !x.IsEmpty);

        public IParagraph BriefDescription { get; set; } = null!;
        public List<IParagraph> DetailedDescription { get; } = new();

        internal void MergeWith(Descriptions other)
        {
            if (this.BriefDescription.IsEmpty)
            {
                this.BriefDescription = other.BriefDescription;
            }
            if (this.DetailedDescription.Count == 0)
            {
                this.DetailedDescription.AddRange(other.DetailedDescription);
            }
        }
    }
}
