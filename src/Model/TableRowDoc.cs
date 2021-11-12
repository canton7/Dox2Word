using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class TableRowDoc
    {
        public List<TableCellDoc> Cells { get; } = new();
    }
}