using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
namespace MyExcelMAUIApp;

public class Table
{
    [JsonPropertyName("cells")]
    [JsonInclude]
    public List<Cell> cells;
    [JsonPropertyName("CurrentCountColumn")]
    public int CurrentCountColumn { get; set; }
    [JsonPropertyName("CurrentCountRoW")]
    public int CurrentCountRow { get; set; }

    public Table(int row, int col)
    {
        this.cells = new List<Cell>();
        this.CurrentCountRow = row;
        this.CurrentCountColumn = col;
    }

    [JsonConstructor]
    public Table(int currentCountRow, int currentCountColumn, List<Cell> cells)
    {
        this.cells = cells;
        this.CurrentCountRow = currentCountRow;
        this.CurrentCountColumn = currentCountColumn;
    }

    public void AddCellToTable(Cell cell)
    {
        cells.Add(cell);
    }

    public void ClearTable(){
        cells.Clear();
    }

    public Cell FindCellByEntry(int row, int column)
    {
        Cell entryCell = null;
        foreach (Cell cell in cells)
        {
            if (cell.cellRow == row && cell.cellColumn == column)
            {
                entryCell = cell;
            }
        }
        return entryCell;
    }

    public Cell FindCellByName(string name)
    {
        Cell nameCell = null;
        foreach (Cell cell in cells)
        {
            if (cell.cellName == name)
            {
                nameCell = cell;
            }
        }
        return nameCell;
    }

    public bool TryCalculate(Cell ourcell)
    {
        var usedcells = ParseName(ourcell.expression);
        ourcell.dependences = new List<string>();
        foreach (string celln in usedcells)
        {
            ourcell.dependences.Add(celln);
            Cell cellByName = FindCellByName(celln);
            cellByName.appearance.Add(ourcell.cellName);
        }
        if (CheckRecursion(ourcell.cellName, ourcell))
        {
            return false;
        }
        return true;
    }

    public static List<string> ParseName(string expression)
    {
        string[] lst = expression.Split(new char[]{'.', ',', ' ', '(', ')', '-', '+', '^', '*', '/'},
        StringSplitOptions.RemoveEmptyEntries);
        List <string> ans = new List<string>();
        foreach(var str in lst)
        {
            if( str[0] >= 'A' && str[0] <= 'Z' && str[str.Length-1] >= '0' && str[str.Length-1] <= '9' )
            {
                ans.Add(str);
            }
        }
        return ans;
    }

    public bool CheckRecursion(string target, Cell cell)
    {
        bool ans = false;
        foreach (string cellt in cell.dependences)
        {
            if (cellt == target)
            {
                ans = true;
            }
            else
            {
                Cell ourcell = FindCellByName(cellt);
                ans |= CheckRecursion(target, ourcell);
            }
        }
        return ans;
    }

    public void RecalculateRecursively(Cell cell)
    {
        foreach (string cellt in cell.appearance)
        {
            Cell ourcell = FindCellByName(cellt);
            Console.WriteLine(cellt);
            ourcell.Calculate();
            RecalculateRecursively(ourcell);
        }
    }   
}