using GoogeWrapperLibrary.Data.Models;
using GoogeWrapperLibrary.Services;
using Google.Apis.Sheets.v4.Data;
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
        public void GetSheet(string sheetNameId)
        {
            if (_sheetService == null) return;
            Sheet = _sheetService.GetSheet(sheetNameId);
            if (Sheet != null)
            {
                SpreadSheetId = sheetNameId;
            }
            else SpreadSheetId = null;
        }
        public bool SetCredentialsDirectory(string credentialsDirectory)
        {
            if (!string.IsNullOrEmpty(credentialsDirectory))
            {
                _sheetService = new SheetService(credentialsDirectory);
                return true;
            }
            else return false;
        }
    }
}
