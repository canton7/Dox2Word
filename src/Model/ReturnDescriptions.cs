using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class ReturnDescriptions
    {
        public List<IParagraph> Description { get; } = new();
        public List<ReturnValueDoc> Values { get; } = new();
    }
}
