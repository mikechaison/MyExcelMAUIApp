using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace MyExcelMAUIApp;
public static class SavesManager 
{
    public async static void SaveToJsonTable(Table table, string path)
    {
        string JTable = JsonSerializer.Serialize(table);
        File.WriteAllText(@path, JTable);
    }

    public static Table ReadJsonTable(string path)
    {
        string JTable = File.ReadAllText(@path);
        Table newtable = JsonSerializer.Deserialize<Table>(JTable, new JsonSerializerOptions 
        {
            IncludeFields = true,
            PropertyNameCaseInsensitive = true
        });
        return newtable;
    }

}