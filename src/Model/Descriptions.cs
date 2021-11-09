using System.Collections.Generic;
using System.Linq;

namespace Dox2Word.Model
{
    public class Descriptions
    {
        public bool HasDescriptions => this.BriefDescription.Count > 0 || this.DetailedDescription.Any(x => x.Count > 0);

        public IParagraph BriefDescription { get; set; } = null!;
        public List<IParagraph> DetailedDescription { get; } = new();

        internal void MergeWith(Descriptions other)
        {
            if (this.BriefDescription.Count == 0)
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
