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
            if (Sheet != null)
            {
                SpreadSheetId = sheetName;
            }
            else SpreadSheetId = null;
        }
        public string SpreadSheetId { get; set; }
        public SheetModel Sheet { get; set; }
        public Cell GetCell(int rowIndex, int columnIndex) => _sheetService.GetCell($"{SpreadSheetId}/{Sheet.SheetId}", columnIndex, rowIndex);
        public Cell GetCell(int rowIndex, string columnLetter) => _sheetService.GetCell($"{SpreadSheetId}/{Sheet.SheetId}", rowIndex, columnLetter);
        public Row GetRow(int rowIndex) => _sheetService.GetRow($"{SpreadSheetId}/{Sheet.SheetId}", rowIndex);
        public Column GetColumn(int columnIndex) => _sheetService.GetColumn($"{SpreadSheetId}/{Sheet.SheetId}", columnIndex);
        public Column GetColumn(string columnLetter) => _sheetService.GetColumn($"{SpreadSheetId}/{Sheet.SheetId}", columnLetter);
        public Cell PutCell(int rowIndex, int columnIndex, string Value) => _sheetService.PutCell($"{SpreadSheetId}/{Sheet.SheetId}", rowIndex, columnIndex, Value);
        public Cell PutCell(int rowIndex, int columnIndex, Cell Value) => _sheetService.PutCell($"{SpreadSheetId}/{Sheet.SheetId}", rowIndex, columnIndex, Value);
        public Row PutRow(int rowIndex,  Row value) => _sheetService.PutRow($"{SpreadSheetId}/{Sheet.SheetId}", rowIndex, value);
        public Column PutColumn(int columnIndex, Column columnValue) => _sheetService.PutColumn($"{SpreadSheetId}/{Sheet.SheetId}", columnIndex, columnValue);
    }
}
