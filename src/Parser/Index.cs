using System.Collections.Generic;
using System.Linq;
using Dox2Word.Parser.Models;

namespace Dox2Word.Parser
{
    public class Index
    {
        public Dictionary<string, Compound> Compounds { get; }

        public Index(DoxygenIndex doxygenIndex)
        {
            this.Compounds = doxygenIndex.Compounds.ToDictionary(x => x.RefId, x => x);
        }

        // Compounds doesn't contain members, and we do want to link to members. We just care about not linking to files
        public bool ShouldReference(string refId) => !this.Compounds.TryGetValue(refId, out var compound) || compound.Kind != CompoundKind.File;

    }
}
