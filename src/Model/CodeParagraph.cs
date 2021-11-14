using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class CodeParagraph : IParagraph
    {
        public List<string> Lines { get; } = new();

        public bool IsEmpty => this.Lines.Count == 0;

        public void TrimTrailingWhitespace() { }
    }
}
