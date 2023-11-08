using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
namespace MyExcelMAUIApp;

public class Table
{
    [JsonPropertyName("Cells")]
    [JsonInclude]
    public List<Cell> Cells { get; set; }
    [JsonPropertyName("CurrentCountColumn")]
    public int CurrentCountColumn { get; set; }
    [JsonPropertyName("CurrentCountRoW")]
    public int CurrentCountRow { get; set; }
    public Dictionary<string, Cell> CellNames { get; set; } = new Dictionary<string, Cell>(); //для отримання клітин за іменами

    public Table(int row, int col)
    {
        this.Cells = new List<Cell>();
        this.CurrentCountRow = row;
        this.CurrentCountColumn = col;
    }

    [JsonConstructor]
    public Table(int currentCountRow, int currentCountColumn, List<Cell> cells)
    {
        this.Cells = cells;
        this.CurrentCountRow = currentCountRow;
        this.CurrentCountColumn = currentCountColumn;
        foreach(Cell OurCell in this.Cells)
        {
            this.CellNames[OurCell.CellName] = OurCell;
        }
    }

    //Додавання клітини в таблицю
    public void AddCellToTable(Cell Cll)
    {
        Cells.Add(Cll);
    }

    //Очищення усієї таблиці
    public void ClearTable(){
        Cells.Clear();
    }

    //Очищення даних клітини таблиці
    public void ClearCellData(Cell OurCell)
    {
        foreach (string Celln in OurCell.Dependences)
        {
            Cell CellByName = CellNames[Celln];
            CellByName.Appearance.Remove(OurCell.CellName);
        }
        OurCell.Expression = "";
        OurCell.Dependences = new List<string>();
        Calculator.GlobalScope[OurCell.CellName] = 0;
        OurCell.CellEntry.Text = "";
            
    }

    //Підготовка до перевірки на циклічну залежність
    public bool TryCalculate(Cell OurCell, string Expr)
    {
        var PrevCells = ParseName(OurCell.Expression);
        foreach (string Celln in PrevCells)
        {
            Cell CellByName = CellNames[Celln];
            CellByName.Appearance.Remove(OurCell.CellName);
        }
        var UsedCells = ParseName(Expr);
        OurCell.Dependences = new List<string>();
        foreach (string Celln in UsedCells)
        {
            OurCell.Dependences.Add(Celln);
            Cell CellByName = CellNames[Celln];
            CellByName.Appearance.Add(OurCell.CellName);
        }
        OurCell.Expression = Expr;
        if (CheckRecursion(OurCell.CellName, OurCell))
        {
            ClearCellData(OurCell);
            return false;
        }
        return true;
    }

    //Добування ідентифікаторів з виразу в клітині
    public static List<string> ParseName(string Expression)
    {
        string[] Lst = Expression.Split(new char[]{'.', ',', ' ', '(', ')', '-', '+', '^', '*', '/'},
        StringSplitOptions.RemoveEmptyEntries);
        List <string> Ans = new List<string>();
        foreach(var Str in Lst)
        {
            if( Str[0] >= 'A' && Str[0] <= 'Z' && Str[Str.Length-1] >= '0' && Str[Str.Length-1] <= '9' )
            {
                Ans.Add(Str);
            }
        }
        return Ans;
    }

    //Перевірка на циклічну залежність
    public bool CheckRecursion(string Target, Cell Cll)
    {
        bool Ans = false;
        foreach (string Cellt in Cll.Dependences)
        {
            if (Cellt == Target)
            {
                Ans = true;
            }
            else
            {
                Cell OurCell = CellNames[Cellt];
                Ans |= CheckRecursion(Target, OurCell);
            }
        }
        return Ans;
    }

    //Переобчислення виразів в клітинах, які залежать від обраної клітини
    public void RecalculateRecursively(Cell Cll)
    {
        foreach (string Cellt in Cll.Appearance)
        {
            Cell OurCell = CellNames[Cellt];
            Console.WriteLine(Cellt);
            OurCell.Calculate();
            RecalculateRecursively(OurCell);
        }
    }   
}