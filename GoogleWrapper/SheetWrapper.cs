using GoogeWrapperLibrary;
using GoogeWrapperLibrary.Data.Models;
using GoogeWrapperLibrary.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace GoogleWrapper
{
    static class Program
    {
        private static SheetWrapper _wrapper;
        private static string exit = "exit";
        static void Main(string[] args)
        {
            _wrapper = new SheetWrapper();
            Console.WriteLine("Application is ready to work");
            while (true)
            {
                var requested = Console.ReadLine().Split(' ');
                if (requested[0].Equals(exit)) break;
                if (requested[0] == "wrapper" && requested.Length>1)
                {
                    if (requested[1] == "initialize")
                    {
                        _wrapper.GetSheet(requested[2]);
                        if (_wrapper.Sheet != null)
                        {
                            Console.WriteLine($"Sheet {_wrapper.Sheet.SheetId} was succesfully download\nSheet has {_wrapper.Sheet.Rows.Count} rows");
                        }
                        else
                        {
                            Console.WriteLine("SpreadSheet does not exist or user credential was not setted to application. Please, try another spreadsheet id for initialize or set credentials by command 'set'[argument=credentials directory].");continue;
                        }
                    }
                    if(requested[1] == "set")
                    {
                        if(_wrapper.SetCredentialsDirectory(requested[2]))
                            Console.WriteLine("Credentials was succesfully setted");
                        else Console.WriteLine("Incorrectly directory, please make sure that path is actual or exists");
                    }
                    else
                    {
                        if (_wrapper == null)
                        {
                            Console.WriteLine("Application did not initialize any spreadsheet.Pleas initialize application by keyword 'initialize' and argument (spreadsheet id)"); continue;
                        }
                        if (_wrapper.SpreadSheetId == null)
                        { 
                            Console.WriteLine("In application was downloaded unexisting spreadsheet. Reinitialize application by existing id of spreadsheet."); continue;
                        }
                        switch (requested[1])
                        {
                            case "get":
                                {
                                    if (requested.Length >= 3)
                                    {
                                        switch (requested[2])
                                        {
                                            case "cell":
                                                {
                                                    if (requested.Length == 3)
                                                    { 
                                                        Console.WriteLine("Enter arguments for getting cell: index row and index(letter) column separated by space.");break; 
                                                    }
                                                    if (int.TryParse(requested[3], out int indexRow))
                                                    {
                                                        if (int.TryParse(requested[4], out int columnIndex))
                                                        {
                                                            var cell = _wrapper.GetCell(indexRow, columnIndex);
                                                            if (cell != null)
                                                                Console.WriteLine($"Requested cell has value: {cell.Value}\nCell from row {cell.RowIndex} and column {cell.ColumnLetter}");
                                                            else Console.WriteLine("Cell does not exist in this spreadsheet");
                                                        }
                                                        else
                                                        {
                                                            if (requested[4].Length == 1)
                                                            {
                                                                var cell = _wrapper.GetCell(indexRow, requested[4]);
                                                                if (cell != null)
                                                                    Console.WriteLine($"Requested cell has value: {cell.Value}\nCell from row {cell.RowIndex} and column {cell.ColumnLetter}");
                                                                else Console.WriteLine("Cell does not exist in this spreadsheet");
                                                            }
                                                            if (requested[4].Length > 1)
                                                                Console.WriteLine("Incorrect argument. Please, enter a number or single char for 'A' to 'Z'");
                                                            if (requested[4].Length < 1)
                                                                Console.WriteLine("Please, check that you`ve entered argument after keyword 'cell'");
                                                        }
                                                    }
                                                    break;
                                                }
                                            case "row":
                                                {
                                                    if (requested.Length == 3)
                                                    {
                                                        Console.WriteLine("Enter arguments for getting row: index row."); break;
                                                    }
                                                    if (int.TryParse(requested[3], out int indexRow))
                                                    {
                                                        var row = _wrapper.GetRow(indexRow);
                                                        if (row != null)
                                                        {
                                                            Console.WriteLine($"Requested row {row.Index} has {row.Cells.Count} cells.\n Cells are: ");
                                                            foreach (var cell in row.Cells)
                                                            {
                                                                Console.Write($"{cell.Value} ");
                                                            }
                                                            Console.WriteLine();
                                                        }
                                                        else Console.WriteLine("Row does not exist. Make sure that you`ve entered correctly index from spreadsheet");
                                                    }
                                                    break;
                                                }
                                            case "column":
                                                {
                                                    if (requested[3] == "")
                                                    {
                                                        Console.WriteLine("Enter arguments for getting column: index row and index(letter) column separated by space."); break;
                                                    }
                                                    if (int.TryParse(requested[3], out int indexColumn))
                                                    {
                                                        var column = _wrapper.GetColumn(indexColumn);
                                                        if (column != null)
                                                        {
                                                            Console.WriteLine($"Requested row {column.Index} has {column.Cells.Count} cells.\n Cells are: ");
                                                            foreach (var cell in column.Cells)
                                                            {
                                                                Console.Write($"{cell.Value} ");
                                                            }
                                                            Console.WriteLine();
                                                        }
                                                        else Console.WriteLine("Column does not exist. Make sure that you`ve entered correctly index from spreadsheet");
                                                    }
                                                    else
                                                    {
                                                        if (requested[3].Length == 1)
                                                        {
                                                            var column = _wrapper.GetColumn(requested[3]);
                                                            if (column != null)
                                                            {
                                                                Console.WriteLine($"Requested column {column.Index} has {column.Cells.Count} cells.\n Cells are: ");
                                                                foreach (var cell in column.Cells)
                                                                {
                                                                    Console.Write($"{cell.Value}\n");
                                                                }
                                                                Console.WriteLine(); break;
                                                            }
                                                            else Console.WriteLine("Column does not exist. Make sure that you`ve entered correctly index from spreadsheet"); break;
                                                        }
                                                        else
                                                        {
                                                            if (requested[4].Length > 1)
                                                            {
                                                                Console.WriteLine("Incorrect argument. Please, enter a number or single char for 'A' to 'Z'");break;
                                                            }
                                                            if (requested[4].Length < 1)
                                                            {
                                                                Console.WriteLine("Please, check that you`ve entered argument after keyword 'column'");break;
                                                            }
                                                        }
                                                    }
                                                    break;
                                                }
                                            default: Console.WriteLine("Incorrect argument. Please, enter the existing entities like 'row', 'column' or 'cell'"); break;
                                        }
                                    }
                                    else Console.WriteLine("Incorrect argument. Please, enter the existing entities like 'row', 'column' or 'cell'");
                                    break;
                                }
                            case "put":
                                {
                                    if (requested.Length >= 3)
                                    {
                                        switch (requested[2])
                                        {
                                            case "cell":
                                                {
                                                    if (requested.Length == 3)
                                                    {
                                                        Console.WriteLine("Enter arguments for updating cell: index row and index(letter) column separated by space."); break;
                                                    }
                                                    if (int.TryParse(requested[3], out int indexRow))
                                                    {
                                                        if (int.TryParse(requested[4], out int columnIndex))
                                                        {
                                                            var value = requested[5];
                                                            var updatedCell = _wrapper.PutCell(indexRow, columnIndex, value);
                                                            if (updatedCell != null)
                                                            {
                                                                Console.WriteLine($"Cell {updatedCell.ColumnLetter}{updatedCell.RowIndex} was updated.\nNew values is {updatedCell.Value}");
                                                            }
                                                            else
                                                            {
                                                                Console.WriteLine("Bad update request, please try again.");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (requested[4].Length == 1)
                                                            {
                                                                var cell = _wrapper.GetCell(indexRow, requested[4]);
                                                                if (cell != null)
                                                                    Console.WriteLine($"Requested cell has value: {cell.Value}\nCell from row {cell.RowIndex} and column {cell.ColumnLetter}");
                                                                else Console.WriteLine("Cell does not exist in this spreadsheet");
                                                            }
                                                            if (requested[4].Length > 1)
                                                                Console.WriteLine("Incorrect argument. Please, enter a number or single char for 'A' to 'Z'");
                                                            if (requested[4].Length < 1)
                                                                Console.WriteLine("Please, check that you`ve entered argument after keyword 'cell'");
                                                        }
                                                    }
                                                    break;
                                                }
                                            case "row":
                                                {
                                                    if (requested.Length == 3)
                                                    {
                                                        Console.WriteLine("Enter arguments for updating row: index row."); break;
                                                    }
                                                    if (int.TryParse(requested[3], out int indexRow))
                                                    {
                                                        var subStringRequested = requested.Skip(4).ToArray();
                                                        var updatedRow = new Row();
                                                        updatedRow.Index = indexRow;
                                                        updatedRow.Cells = new List<Cell>();
                                                        for (int i = 0; i < subStringRequested.Length; i++)
                                                        {
                                                            var cell = new Cell();
                                                            cell.Value = subStringRequested[i];
                                                            cell.RowIndex = indexRow;
                                                            cell.ColumnIndex = i + 1;
                                                            cell.ColumnLetter = Mapper.GetColumnLetter(cell.ColumnIndex);
                                                            updatedRow.Cells.Add(cell);
                                                        }
                                                        var resultRow = _wrapper.PutRow(indexRow, updatedRow);
                                                        if (resultRow != null)
                                                        {
                                                            Console.WriteLine($"Row {resultRow.Index} was updated and has new cells` values :");
                                                            foreach (var cell in resultRow.Cells)
                                                            {
                                                                Console.WriteLine(cell.Value);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Console.WriteLine("Requested row isn`t existing or updated row has incorrect input data.");
                                                        }
                                                    }
                                                    break;
                                                }
                                            case "column":
                                                {
                                                    if (requested.Length == 3)
                                                    {
                                                        Console.WriteLine("Enter arguments for updating column: index column and updated cell`s values for column separated by space."); break;
                                                    }
                                                    if (requested[3] == "")
                                                    {
                                                        Console.WriteLine("Here is no argument. Please, enter at least one value for updating");break;
                                                    }
                                                    if (int.TryParse(requested[3], out int columnIndex))
                                                    {
                                                        var subStringRequested = requested.Skip(4).ToArray();
                                                        var updatedColumn = new Column();
                                                        updatedColumn.Index = columnIndex;
                                                        updatedColumn.Cells = new List<Cell>();
                                                        for (int i = 0; i < subStringRequested.Length; i++)
                                                        {
                                                            var cell = new Cell();
                                                            cell.Value = subStringRequested[i];
                                                            cell.ColumnIndex = columnIndex;
                                                            cell.RowIndex = i + 1;
                                                            cell.ColumnLetter = Mapper.GetColumnLetter(cell.ColumnIndex);
                                                            updatedColumn.Cells.Add(cell);
                                                        }
                                                        var resultColumn = _wrapper.PutColumn(columnIndex,updatedColumn);
                                                        if (resultColumn != null)
                                                        {
                                                            Console.WriteLine($"Column {resultColumn.Name} was updated and has new cells` values :");
                                                            foreach (var cell in resultColumn.Cells)
                                                            {
                                                                Console.WriteLine(cell.Value);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            Console.WriteLine("Requested row isn`t existing or updated row has incorrect input data.");
                                                        }
                                                    }
                                                    break;
                                                }
                                            default: Console.WriteLine("Incorrect argument. Please, enter the existing entities like 'row', 'column' or 'cell'"); break;
                                        }
                                    }
                                    else Console.WriteLine("Incorrect argument. Please, enter the existing entities like 'row', 'column' or 'cell'");
                                    break;
                                }
                        }
                    }
                }
                     
                
                else
                {
                    Console.WriteLine("Incorrect input data. Please, make sure that you`ve wrote 'wrapper' or 'exit' keywords.");
                }
                               
            }
        }
    }
}

