using Dox2Word.Parser.Models;
using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class CellDoc
    {
        public bool IsHeading { get; set; }
        public int ColSpan { get; set; }
        public int RowSpan { get; set; }

        public List<IParagraph> Paragraphs { get; } = new();
    }
}