using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
namespace MyExcelMAUIApp;

public class Cell
{
    [JsonPropertyName("cellRow")]
    public int cellRow { get; set; }
    [JsonPropertyName("cellColumn")]
    public int cellColumn { get; set; }
    [JsonPropertyName("cellName")]
    public string cellName { get; set; }
    [JsonIgnore]
    public Entry cellEntry { get; set; }
    [JsonPropertyName("expression")]
    public string expression { get; set; }
    [JsonPropertyName("appearance")]
    [JsonInclude]
    public List <string> appearance { get; set; }
    [JsonPropertyName("dependences")]
    [JsonInclude]
    public List <string> dependences { get; set; }

    public Cell(int row, int column, Entry entry, string name)
    {
        this.cellRow = row;
        this.cellColumn = column;
        this.cellEntry = entry;
        this.expression = "";
        this.cellName = name;
        this.appearance = new List<string>();
        this.dependences = new List<string>();
    }

    [JsonConstructor]
    public Cell(int cellRow, int cellColumn, string expression, string cellName, List<string> appearance, List<string> dependences)
    {
        this.cellRow = cellRow;
        this.cellColumn = cellColumn;
        this.expression = expression;
        this.cellName = cellName;
        this.appearance = appearance;
        this.dependences = dependences;
    }

    public void Calculate()
    {
        if (this.expression != "")
        {
            var val  = Calculator.Evaluate(this.expression);
            var content = val.ToString();
            Calculator.GlobalScope[this.cellName] = val;
            this.cellEntry.Text = content;
        }
    }
}