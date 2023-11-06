using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
namespace MyExcelMAUIApp;

public class Table
{
    [JsonPropertyName("cells")]
    [JsonInclude]
    public List<Cell> cells { get; set; }
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

    //Додавання клітини в таблицю
    public void AddCellToTable(Cell cell)
    {
        cells.Add(cell);
    }

    //Очищення усієї таблиці
    public void ClearTable(){
        cells.Clear();
    }

    //Знаходження клітини за його місцезнаходженням
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

    //Знаходження клітини за його іменем
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

    //Очищення даних клітини таблиці
    public void ClearCellData(Cell ourcell)
    {
        foreach (string celln in ourcell.dependences)
        {
            Cell cellByName = FindCellByName(celln);
            cellByName.appearance.Remove(ourcell.cellName);
        }
        ourcell.expression = "";
        ourcell.dependences = new List<string>();
        Calculator.GlobalScope[ourcell.cellName] = 0;
        ourcell.cellEntry.Text = "";
            
    }

    //Підготовка до перевірки на циклічну залежність
    public bool TryCalculate(Cell ourcell, string expr)
    {
        var prevcells = ParseName(ourcell.expression);
        foreach (string celln in prevcells)
        {
            Cell cellByName = FindCellByName(celln);
            cellByName.appearance.Remove(ourcell.cellName);
        }
        var usedcells = ParseName(expr);
        ourcell.dependences = new List<string>();
        foreach (string celln in usedcells)
        {
            ourcell.dependences.Add(celln);
            Cell cellByName = FindCellByName(celln);
            cellByName.appearance.Add(ourcell.cellName);
        }
        ourcell.expression = expr;
        if (CheckRecursion(ourcell.cellName, ourcell))
        {
            ClearCellData(ourcell);
            return false;
        }
        return true;
    }

    //Добування ідентифікаторів з виразу в клітині
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

    //Перевірка на циклічну залежність
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

    //Переобчислення виразів в клітинах, які залежать від обраної клітини
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