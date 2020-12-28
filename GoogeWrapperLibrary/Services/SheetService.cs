using GoogeWrapperLibrary.Data;
using GoogeWrapperLibrary.Data.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace GoogeWrapperLibrary.Services
{
    public class SheetService
    {
        private GoogleCredential _currentUser;
        private SheetsService _service;
        public SheetService(string credentialsDirectory)
        {
            _currentUser = GetCurrentUser(credentialsDirectory).User;
            _service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = _currentUser
            });
            Sheets = new List<SheetModel>();
        }
        public List<SheetModel> Sheets { get; }

        public Cell GetCell(string sheetName, int columnIndex, int rowIndex)
        {
            var requestIds = sheetName.Split('/');
            var spreadSheetId = requestIds[0];
            var sheetId = requestIds[1];
            var sheet = Sheets.FirstOrDefault(s => s.SheetId.ToString() == sheetId);
            return sheet.Rows[rowIndex-1].Cells[columnIndex - 1];
        }
        public Cell GetCell(string sheetName, int rowIndex, string columnLetter)
        {
            var requestIds = sheetName.Split('/');
            var spreadSheetId = requestIds[0];
            var sheetId = requestIds[1];
            var sheet = Sheets.FirstOrDefault(s => s.SheetId.ToString() == sheetId);
            return sheet.Columns.FirstOrDefault(t => t.Name == columnLetter).Cells[rowIndex - 1];
        }
        public Row GetRow(string sheetName, int rowIndex)
        {
            var requestIds = sheetName.Split('/');
            var spreadSheetId = requestIds[0];
            var sheetId = requestIds[1];
            return Sheets.FirstOrDefault(s => s.SheetId.ToString() == sheetId).Rows.FirstOrDefault(r => r.Index == rowIndex);
        }
        public Column GetColumn(string sheetName,string columnLetter)
        {
            var requestIds = sheetName.Split('/');
            var spreadSheetId = requestIds[0];
            var sheetId = requestIds[1];
            return Sheets.FirstOrDefault(s => s.SheetId.ToString() == sheetId).Columns.FirstOrDefault(c => c.Name == columnLetter);
        }
        public Column GetColumn(string sheetName, int columnIndex)
        {
            var requestIds = sheetName.Split('/');
            var spreadSheetId = requestIds[0];
            var sheetId = requestIds[1];
            return Sheets.FirstOrDefault(s => s.SheetId.ToString() == sheetId).Columns.FirstOrDefault(c => c.Index == columnIndex);
        }
        public SheetModel GetSheet(string sheetName)
        {
            var request = _service.Spreadsheets.Get(sheetName);
            request.IncludeGridData = true;
            Spreadsheet spreadSheet;
            try
            {
                 spreadSheet = request.Execute();
            }
            catch (System.Exception)
            {
                return null;
            }
          
            foreach (var sheet in spreadSheet.Sheets)
            {
                Sheets.Add(Mapper.MapToSheetModel(sheet, sheet.Properties.SheetId ?? 0));
            }
            return Sheets.FirstOrDefault(s => s.SheetId.ToString() == sheetName) ?? Sheets[0];
           
        }
        public Cell PutCell(string sheetName, int rowIndex, int columnIndex, string cellValue)
        {
            var requestIds = sheetName.Split('/');
            var spreadSheetId = requestIds[0];
            var sheetId = requestIds[1];
            var rowDTO = Sheets.FirstOrDefault(s => s.SheetId.ToString() == sheetId).Rows[rowIndex - 1];
            var cellDTO = rowDTO.Cells[columnIndex - 1];
            rowDTO.Cells.Remove(cellDTO);
            cellDTO.Value = cellValue;
            rowDTO.Cells.Insert(columnIndex - 1, cellDTO);
            var range= $"{cellDTO.ColumnLetter}{cellDTO.RowIndex}";
            var updateRequest = CreateUpdateRequest(cellDTO.Value, spreadSheetId, range,"ROWS", SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED);
            var response = updateRequest.Execute();
            if (response.UpdatedCells == 1)
                return cellDTO;
            else return null;
        }
        public Cell PutCell(string sheetName, int rowIndex, int columnIndex, Cell cellValue)
        {
            var requestIds = sheetName.Split('/');
            var spreadSheetId = requestIds[0];
            var sheetId = requestIds[1];
            var cellDTO = Sheets.FirstOrDefault(s => s.SheetId.ToString() == sheetId).Rows[rowIndex - 1].Cells[columnIndex - 1];
            cellDTO = cellValue;
            var range = $"{cellDTO.ColumnLetter}{cellDTO.RowIndex}";
            var updateRequest = CreateUpdateRequest(cellDTO.Value, spreadSheetId, range,"ROWS", SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED);
            var response = updateRequest.Execute();
            if (response.UpdatedCells == 1)
            {
                return cellDTO;
            }
            else return null;
        }
        public Row PutRow(string sheetName, int rowIndex, Row rowValue)
        {
            var requestIds = sheetName.Split('/');
            var spreadSheetId = requestIds[0];
            var sheetId = requestIds[1];
            var rowDTO = Sheets.FirstOrDefault(s => s.SheetId.ToString() == sheetId).Rows[rowIndex - 1];
            rowDTO = rowValue;
            var range = $"{rowDTO.Cells[0].ColumnLetter}{rowDTO.Cells[0].RowIndex}:{rowDTO.Cells.Last().ColumnLetter}{rowDTO.Cells.Last().RowIndex}";
            var request = CreateUpdateRequest(rowDTO.Cells.Select(t=>t.Value), spreadSheetId, range,"ROWS", SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED);
            var response = request.Execute();
            if (response.UpdatedRange == $"'Лист1'!{range}")
                return rowDTO;
            else return null;
        }
        public Column PutColumn(string sheetName, int columnIndex, Column colValue)
        {
            var requestIds = sheetName.Split('/');
            var spreadSheetId = requestIds[0];
            var sheetId = requestIds[1];
            var column = Sheets.FirstOrDefault(s => s.SheetId.ToString() == sheetId).Columns.FirstOrDefault(c => c.Index == columnIndex);
            column = colValue;
            var range = $"{column.Cells[0].ColumnLetter}{column.Cells[0].RowIndex}:{column.Cells.Last().ColumnLetter}{column.Cells.Last().RowIndex}";
            var request = CreateUpdateRequest(column.Cells.Select(c => c.Value), spreadSheetId, range,"COLUMNS", SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED);
            var response = request.Execute();
            if (response.UpdatedRange == $"'Лист1'!{range}")
                return column;
            else return null;
        }
        private SpreadsheetsResource.ValuesResource.UpdateRequest CreateUpdateRequest(object value, 
                                                                                      string spreadSheetIdd, 
                                                                                      string range,
                                                                                      string majorDimension,
                                                                                      SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum option)
        {
            var valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>>();
            var values = new List<object>();
            if(value is IEnumerable<object>)
            {
                foreach (var item in value as IEnumerable<object>)
                {
                    values.Add(item);
                }
            }
            else { values.Add(value); }
            valueRange.Values.Add(values);
            valueRange.Range = range;
            valueRange.MajorDimension = majorDimension;
            var updateRequest = _service.Spreadsheets.Values.Update(valueRange, spreadSheetIdd, valueRange.Range);
            updateRequest.ValueInputOption = option;
            return updateRequest;
        }
        private UserInfo GetCurrentUser(string credentialsDirectory)
        {
            try
            {
                using (var stream = new FileStream(credentialsDirectory, FileMode.Open, FileAccess.Read))
                {
                    var userInfo = new UserInfo();
                    userInfo.User = GoogleCredential.FromStream(stream).CreateScoped(UserInfo.Scopes);
                    return userInfo;
                }
            }
            catch (DirectoryNotFoundException)
            {
                return null;
            }
            
        }
    }
}
