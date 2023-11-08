using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace MyExcelMAUIApp;
public static class SavesManager 
{
    //Збереження json-таблиці
    public async static void SaveToJsonTable(Table OurTable, string Path)
    {
        string JTable = JsonSerializer.Serialize(OurTable);
        File.WriteAllText(@Path, JTable);
    }

    //Зчитування json-таблиці
    public static Table ReadJsonTable(string Path)
    {
        string JTable = File.ReadAllText(@Path);
        Table NewTable = JsonSerializer.Deserialize<Table>(JTable, new JsonSerializerOptions 
        {
            IncludeFields = true,
            PropertyNameCaseInsensitive = true
        });
        return NewTable;
    }
}