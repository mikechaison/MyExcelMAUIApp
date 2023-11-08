using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
namespace MyExcelMAUIApp;

public class Cell
{
    [JsonPropertyName("CellRow")]
    public int CellRow { get; set; }
    [JsonPropertyName("CellColumn")]
    public int CellColumn { get; set; }
    [JsonPropertyName("CellName")]
    public string CellName { get; set; }
    [JsonIgnore]
    public Entry CellEntry { get; set; }
    [JsonPropertyName("Expression")]
    public string Expression { get; set; }
    [JsonPropertyName("Appearance")]
    [JsonInclude]
    public List <string> Appearance { get; set; }
    [JsonPropertyName("Dependences")]
    [JsonInclude]
    public List <string> Dependences { get; set; }

    public Cell(int row, int column, Entry entry, string name)
    {
        this.CellRow = row;
        this.CellColumn = column;
        this.CellEntry = entry;
        this.Expression = "";
        this.CellName = name;
        this.Appearance = new List<string>();
        this.Dependences = new List<string>();
    }

    [JsonConstructor]
    public Cell(int CellRow, int CellColumn, string Expression, string CellName, List<string> Appearance, List<string> Dependences)
    {
        this.CellRow = CellRow;
        this.CellColumn = CellColumn;
        this.Expression = Expression;
        this.CellName = CellName;
        this.Appearance = Appearance;
        this.Dependences = Dependences;
    }

    //Обчислює вираз, написаний для цієї клітини
    public void Calculate()
    {
        if (this.Expression != "")
        {
            var val  = Calculator.Evaluate(this.Expression);
            var content = val.ToString();
            Calculator.GlobalScope[this.CellName] = val;
            this.CellEntry.Text = content;
        }
    }
}