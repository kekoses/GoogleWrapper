using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogeWrapperLibrary.Data.Models
{
    public class Column
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public List<Cell> Cells { get; set; }
    }
}
