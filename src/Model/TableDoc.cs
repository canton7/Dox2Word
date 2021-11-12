using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class TableDoc : IParagraph
    {
        public int NumColumns { get; set; }

        public bool FirstRowHeader { get; set; }
        public bool FirstColumnHeader { get; set; }

        public List<TableRowDoc> Rows { get; } = new();

        int IParagraph.Count => this.Rows.Count;
    }
}
