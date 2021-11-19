using System.Collections.Generic;

namespace Dox2Word.Model
{
    public enum ParameterDirection
    {
        None,
        In,
        Out,
        InOut,
    }

    public class ParameterDoc
    {
        public string Name { get; set; } = null!;
        public List<TextRun> Type { get; set; } = null!;
        public IParagraph Description { get; set; } = null!;
        public ParameterDirection Direction { get; set; } = ParameterDirection.None;
    }
}