using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class DefinitionListParagraph : IParagraph
    {
        public bool IsEmpty => this.Entries.Count == 0;

        public List<DefinitionListEntry> Entries { get; } = new();

        public void TrimTrailingWhitespace() { }
    }

    public class DefinitionListEntry
    {
        public TextParagraph Term { get; set; } = null!;

        public List<IParagraph> Description { get; } = new();
    }
}
