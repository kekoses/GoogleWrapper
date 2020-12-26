using GoogeWrapperLibrary.Data;
using GoogeWrapperLibrary.Data.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Collections.Generic;
using System.Linq;

namespace GoogeWrapperLibrary.Services
{
    public class SheetService
    {
        private UserCredential _currentUser;
        private SheetsService _service;
        public SheetService()
        {
            _currentUser = UserInfo.User;
            _service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = _currentUser
            });
            Sheets = new List<SheetModel>();
        }
        public List<SheetModel> Sheets { get; }

        public Cell GetCell(string sheetName, int columnIndex, int rowIndex)
        {
            var sheet = Sheets.FirstOrDefault(t => t.Name == sheetName);
            return sheet.Rows[rowIndex-1].Cells[columnIndex - 1];
        }
        public Cell GetCell(string sheetName, int rowIndex, string columnLetter)
        {
            var sheet = Sheets.FirstOrDefault(t => t.Name == sheetName);
            return sheet.Columns.FirstOrDefault(t => t.Name == columnLetter).Cells[rowIndex - 1];
        }
        public Row GetRow(string sheetName, int rowIndex)
        {
            return Sheets.FirstOrDefault(s => s.Name == sheetName).Rows.FirstOrDefault(r => r.Index == rowIndex);
        }
        public SheetModel GetSheet(string sheetName)
        {
            var request = _service.Spreadsheets.Get(sheetName);
            request.IncludeGridData = true;
            var spreadSheet = request.Execute();
            foreach (var sheet in spreadSheet.Sheets)
            {
                Sheets.Add(Mapper.MapToSheetModel(sheet));
            }
            return Sheets[0];
           
        }
        public Cell PutCell(string sheetName, int rowIndex, int columnIndex, string cellValue)
        {
            var cellsDTO = Sheets.FirstOrDefault(s => s.Name == sheetName).Rows[rowIndex - 1].Cells;
            var cellDTO = cellsDTO[columnIndex - 1];
            cellDTO.Value = cellValue;
            var updateCellRequest = new UpdateCellsRequest();
            updateCellRequest.Range = new GridRange();
            updateCellRequest.Rows = new List<RowData> { new RowData { Values = cellsDTO.Select(Mapper.MapToCellData).ToList() } };
            updateCellRequest.Rows[0].Values.Add(Mapper.MapToCellData(cellDTO));
            updateCellRequest.Range.StartColumnIndex = cellDTO.ColumnIndex;
            updateCellRequest.Range.StartRowIndex = cellDTO.RowIndex;
            var spreadSheetUpdateRequest = new BatchUpdateSpreadsheetRequest();
            spreadSheetUpdateRequest.Requests = new List<Request>();
            spreadSheetUpdateRequest.Requests.Add(new Request { UpdateCells = updateCellRequest });
            var request=_service.Spreadsheets.BatchUpdate(spreadSheetUpdateRequest, "1H6uZOYrVOoq7x2WBrk4IPEVpjZEIkvZ6qK_SKaetNb4/id=1904996944");
            request.Execute();
            return cellDTO;
        }
        public Cell PutCell(string sheetName, int rowIndex, int columnIndex, Cell cellValue)
        {
            var cellDTO = Sheets.FirstOrDefault(s => s.Name == sheetName).Rows[rowIndex - 1].Cells[columnIndex - 1];
            cellDTO = cellValue;
            var updateCellRequest = new UpdateCellsRequest();
            updateCellRequest.Range.StartColumnIndex = cellDTO.ColumnIndex;
            updateCellRequest.Range.StartRowIndex = cellDTO.RowIndex;
            CellData value= Mapper.MapToCellData(cellDTO);
            updateCellRequest.Rows[rowIndex - 1].Values.Add(value);
            var spreadSheetUpdateRequest = new BatchUpdateSpreadsheetRequest();
            spreadSheetUpdateRequest.Requests.Add(new Request { UpdateCells = updateCellRequest });
            var request = _service.Spreadsheets.BatchUpdate(spreadSheetUpdateRequest, "1H6uZOYrVOoq7x2WBrk4IPEVpjZEIkvZ6qK_SKaetNb4");
            var response = request.Execute();
            
            return cellDTO;
        }

        public Row PutRow(string sheetName, int rowIndex, Row rowValue)
        {
            var rowDTO = Sheets.FirstOrDefault(t => t.Name == sheetName).Rows[rowIndex - 1];
            rowDTO = rowValue;
            var updateSpreadSheetRequest = new BatchUpdateSpreadsheetRequest();
            var updateCellRequest = new UpdateCellsRequest();
            updateCellRequest.Range.StartColumnIndex = rowDTO.Cells[0].ColumnIndex;
            updateCellRequest.Range.StartRowIndex = rowDTO.Cells[0].RowIndex;
            updateCellRequest.Range.EndColumnIndex = rowDTO.Cells.Last().ColumnIndex;
            updateCellRequest.Range.EndRowIndex = rowDTO.Cells.Last().RowIndex;
            updateCellRequest.Rows.Add(Mapper.MapToRowData(rowDTO));
            updateSpreadSheetRequest.Requests.Add(new Request { UpdateCells = updateCellRequest });
            var result = _service.Spreadsheets.BatchUpdate(updateSpreadSheetRequest, sheetName);
            return rowDTO;
           
        }
    }
}
