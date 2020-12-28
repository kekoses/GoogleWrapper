using System.Collections.Generic;

namespace GoogeWrapperLibrary.Data.Models
{
    public class SheetModel
    {
        public int SheetId { get; set; }
        public List<Cell> Cells { get; set; }
        public List<Row> Rows { get; set; }
        public List<Column> Columns { get; set; }
    }
}
