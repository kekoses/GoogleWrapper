using GoogeWrapperLibrary.Data.Models;
using Google.Apis.Sheets.v4.Data;

namespace GoogeWrapperLibrary.Services
{
    public static class Mapper
    {
        public static Cell MapToCell(CellData data, int columnIndex, int rowIndex)
        {
            return new Cell() { ColumnIndex = columnIndex, RowIndex = rowIndex, Value = data.FormattedValue };
        }
        private const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        public static Row MapToRow(RowData rowData, int rowIndex)
        {
            var row = new Row();
            row.Cells = new System.Collections.Generic.List<Cell>();
            row.Index = rowIndex;
            for (int i = 0; i < rowData.Values.Count; i++)
            {
                var cellDTO = MapToCell(rowData.Values[i], i + 1, rowIndex);
                row.Cells.Add(cellDTO);
                
            }
            return row;
        }
        public static Column MapToColumn(int columnIndex)
        {
            var column = new Column();
            column.Index = columnIndex;
            column.Name = alphabet[columnIndex - 1].ToString();
            column.Cells = new System.Collections.Generic.List<Cell>();
            return column;
        }
        public static Column FillColumn(SheetModel sheetDTO,Column column)
        {
            int columnNumber = 0;
            for (int i = 0; i < sheetDTO.Columns.Count; i++)
            {
                foreach (var row in sheetDTO.Rows)
                {
                    if (row.Cells.Count <= columnNumber )
                        break;
                    var cell = row.Cells[columnNumber];
                    cell.ColumnLetter = column.Name;
                    column.Cells.Add(cell);
                }
                columnNumber++;
            }
            return column;
        }
        public static SheetModel MapToSheetModel(Sheet sheet)
        {
            var sheetDTO = new SheetModel();
            sheetDTO.Rows = new System.Collections.Generic.List<Row>();
            sheetDTO.Columns = new System.Collections.Generic.List<Column>();
            foreach (var gridData in sheet.Data)
            {   
                var rowIndex = 1;
                foreach (var rowData in gridData.RowData)
                {
                    var rowDTO = MapToRow(rowData, rowIndex);
                    if (rowDTO != null)
                    {
                        sheetDTO.Rows.Add(rowDTO);
                        rowIndex++;
                    }
                    else continue;
                }
                var columnIndex = 1;
                foreach (var colData in gridData.ColumnMetadata)
                {
                    sheetDTO.Columns.Add(MapToColumn(columnIndex));
                    columnIndex++;
                }
                for (int i = 0; i < sheetDTO.Columns.Count; i++)
                {
                    var column = sheetDTO.Columns[i];
                    sheetDTO.Columns.Remove(column);
                    var newColumn=FillColumn(sheetDTO,column);
                    sheetDTO.Columns.Insert(i, newColumn);
                }
                
            }
            return sheetDTO;

        }
        public static DimensionProperties MapToColumnMetadata(Column columnDTO)
        {
            var columnDimension = new DimensionProperties();
            columnDimension.DataSourceColumnReference.Name = columnDTO.Name;
            return columnDimension;
        }
        public static Sheet MapToSheet(SheetModel sheetDTO)
        {
            var sheet = new Sheet();
            int columnNumber = 0;
            foreach (var row in sheetDTO.Rows)
            {
                sheet.Data[0].RowData.Add(MapToRowData(row));
                sheet.Data[0].ColumnMetadata.Add(MapToColumnMetadata(sheetDTO.Columns[columnNumber]));
                columnNumber++;
            }
            return sheet;
        }

        public static RowData MapToRowData(Row rowDTO)
        {
            var rowData = new RowData();
            foreach (var cellDTO in rowDTO.Cells)
            {
                rowData.Values.Add(MapToCellData(cellDTO));
            }
            return rowData;
        }

        public static CellData MapToCellData(Cell cellDTO)
        {
            var cellData = new CellData() { FormattedValue = cellDTO.Value };
            return cellData;
        }
    }
}
