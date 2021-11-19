using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class ReturnDescriptions
    {
        public IParagraph Description { get; set; } = null!;
        public List<ReturnValueDoc> Values { get; } = new();
    }
}
