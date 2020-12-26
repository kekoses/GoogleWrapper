using GoogeWrapperLibrary.Data.Models;
using GoogeWrapperLibrary.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogeWrapperLibrary
{
    public class SheetWrapper
    {
        SheetService _sheetService;
        public SheetWrapper(string sheetName)
        {
            _sheetService = new SheetService();
            Sheet = _sheetService.GetSheet(sheetName);
        }
        public SheetModel Sheet { get; set; }
        public Cell GetCell(int rowIndex, int columnIndex) => _sheetService.GetCell(Sheet.Name, columnIndex, rowIndex);
        public Cell GetCell(int rowIndex, string columnLetter) => _sheetService.GetCell(Sheet.Name, rowIndex, columnLetter);
        public Cell PutCell(int rowIndex, int columnIndex, string Value) => _sheetService.PutCell(Sheet.Name, rowIndex, columnIndex, Value);
    }
}
