using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class ReturnValueDoc
    {
        public string Name { get; set; } = null!;
        public List<IParagraph> Description { get; } = new();
    }
}