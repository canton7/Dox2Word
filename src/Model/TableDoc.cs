using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class TableDoc : IParagraph
    {
        public int NumCols { get; set; }

        public List<RowDoc> Rows { get; } = new();

        int IParagraph.Count => this.Rows.Count;
    }
}
