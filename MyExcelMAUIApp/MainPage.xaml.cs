using Microsoft.Maui.Controls;
using Microsoft.Maui.Controls.Compatibility;
using System;
using System.Collections.Generic;
using Grid = Microsoft.Maui.Controls.Grid;
using Microsoft.Maui.Storage;
using CommunityToolkit.Maui.Storage;
using System.Text;

namespace MyExcelMAUIApp
{
    public partial class MainPage : ContentPage
    {
        const int CountColumn = 20; // кількість стовпчиків (A to Z)
        const int CountRow = 50; // кількість рядків
        public Table table;
        public Cell currentcell = null;
        IFileSaver fileSaver;
        CancellationTokenSource cancellationTokenSource = new CancellationTokenSource(); 

        public MainPage(IFileSaver fileSaver)
        {
            table = new Table(CountRow, CountColumn);
            this.fileSaver = fileSaver;
            InitializeComponent();
            CreateGrid();
        }

    //створення таблиці
        private void CreateGrid()
        {
            AddColumnsAndColumnLabels();
            AddRowsAndCellEntries();
        }

        private void ClearGrid()
        {
            while (grid.RowDefinitions.Count > 0)
            {
                int lastRowIndex = grid.RowDefinitions.Count - 1;
                grid.RowDefinitions.RemoveAt(lastRowIndex);
                grid.Children.RemoveAt(lastRowIndex * (table.CurrentCountColumn + 1)); // Remove label
                for (int col = 0; col < table.CurrentCountColumn; col++)
                {
                    grid.Children.RemoveAt((lastRowIndex * table.CurrentCountColumn) + col + 1); // Remove entry
                }
            }
            while (grid.ColumnDefinitions.Count > 0)
            {
                int lastColumnIndex = grid.ColumnDefinitions.Count - 1;
                grid.ColumnDefinitions.RemoveAt(lastColumnIndex);
                grid.Children.RemoveAt(lastColumnIndex); //Remove label
            }
        }

        private void UpdateGrid()
        {
            ClearGrid();
            CreateGrid();
            FillTable();
            CalculateTable();
        }
        private void AddColumnsAndColumnLabels()
        {
            // Додати стовпці та підписи для стовпців
            for (int col = 0; col < table.CurrentCountColumn + 1; col++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition{ Width = 100 });
                if (col > 0)
                {
                    var label = new Label
                    {
                        Text = GetColumnName(col),
                        VerticalOptions = LayoutOptions.Center,
                        HorizontalOptions = LayoutOptions.Center
                    };
                    Grid.SetRow(label, 0);
                    Grid.SetColumn(label, col);
                    grid.Children.Add(label);
                }
                else
                {
                    grid.ColumnDefinitions[col].Width = 40;
                }
            }
        }
        private void AddRowsAndCellEntries()
        {
            // Додати рядки, підписи для рядків та комірки
            for (int row = 0; row < table.CurrentCountRow; row++)
            {
                grid.RowDefinitions.Add(new RowDefinition());
                // Додати підпис для номера рядка
                var label = new Label
                {
                    Text = (row + 1).ToString(),
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center
                };
                Grid.SetRow(label, row + 1);
                Grid.SetColumn(label, 0);
                grid.Children.Add(label);
                // Додати комірки (Entry) для вмісту
                for (int col = 0; col < table.CurrentCountColumn; col++)
                {
                    var entry = new Entry
                    {
                        Text = "",
                        VerticalOptions = LayoutOptions.Center,
                        HorizontalOptions = LayoutOptions.Center,
                        IsReadOnly = true,
                        WidthRequest = 100
                    };
                    entry.Focused += Entry_Focused; // обробник події Focused
                    string cellname = GetColumnName(col + 1)+(row + 1).ToString();
                    var cell = new Cell(row + 1, col + 1, entry, cellname);
                    Calculator.GlobalScope[cellname] = 0;
                    table.AddCellToTable(cell);
                    Grid.SetRow(entry, row + 1);
                    Grid.SetColumn(entry, col + 1);
                    grid.Children.Add(entry);
                }
            }
        }

        private void FillTable()
        {
            foreach (Cell cell in table.cells)
            {
                int ind = table.CurrentCountColumn + (table.CurrentCountColumn + 1) * ( cell.cellRow - 1) + cell.cellColumn;
                Entry cellentry = (Entry)grid.Children.ElementAt(ind);
                cellentry.Text = cell.expression; 
                cell.cellEntry = cellentry;
            }
        }

        private void CalculateTable()
        {
            foreach (Cell cell in table.cells)
            {
                cell.Calculate();
            }
        }

        private string GetColumnName(int colIndex)
        {
            int dividend = colIndex;
            string columnName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }
        // викликається, коли користувач вийде зі зміненої клітинки (втратить фокус)
        private async void Entry_Unfocused(object sender, FocusEventArgs e)
        {
            var entry = (Entry)sender;
            var row = Grid.GetRow(entry) - 1;
            var col = Grid.GetColumn(entry) - 1;
            //entry.Text = textInput.Text;
            //Cell cell = table.FindCellByEntry(row + 1, col + 1);
            //cell.Calculate();
        }
        private void Entry_Focused(object sender, FocusEventArgs e)
        {
            if (currentcell != null)
            {
                currentcell.cellEntry.BackgroundColor = Colors.White;
            }
            var entry = (Entry)sender;
            var row = Grid.GetRow(entry) - 1;
            var col = Grid.GetColumn(entry) - 1;
            currentcell = table.FindCellByEntry(row + 1, col + 1);
            textInput.Text = currentcell.expression;
            cellLabel.Text = currentcell.cellName;
            entry.BackgroundColor = Colors.LightYellow;
        }

        private void TextInput_Return(object sender, FocusEventArgs e)
        {
            if (currentcell != null)
            {
                currentcell.expression = textInput.Text;
                var calc = table.TryCalculate(currentcell);
                if (!calc)
                {
                    textInput.Text = "";
                    currentcell.expression = "";
                    CycleError();
                }
                else{
                    currentcell.Calculate();
                    table.RecalculateRecursively(currentcell);
                }
            }
        }

        private async void CycleError()
        {
            await DisplayAlert("Помилка", "Виявлена циклічна залежність!", "OK");
        }

        private async void SaveButton_Clicked(object sender, EventArgs e)
        {
            
            //string path = await DisplayPromptAsync("Збереження файлу", "Вкажіть шлях та назву файлу");
            using var stream = new MemoryStream(Encoding.Default.GetBytes("Text"));
            var path = await fileSaver.SaveAsync("table.json", stream, cancellationTokenSource.Token);
            SavesManager.SaveToJsonTable(table, path.FilePath);
        }
        
        private async void ReadButton_Clicked(object sender, EventArgs e)
        {
            //string path = await DisplayPromptAsync("Зчитування файлу", "Вкажіть шлях та назву файлу");
            var customFileType = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
            {
                { DevicePlatform.WinUI, new[] { ".json" } }
            });
            var result = await FilePicker.PickAsync(new PickOptions
            {
                PickerTitle = "Оберіть файл",
                FileTypes = customFileType
            });
            table = SavesManager.ReadJsonTable(result.FullPath);
            UpdateGrid();
        }
        private async void ExitButton_Clicked(object sender, EventArgs e)
        {
            bool answer = await DisplayAlert("Підтвердження", "Ви дійсно хочете вийти?", "Так", "Ні");
            if (answer)
            {
                bool answer1 = await DisplayAlert("Збереження", "Зберігати файл?", "Так", "Ні");
                if (answer1)
                {
                    string path = await DisplayPromptAsync("Збереження файлу", "Вкажіть шлях та назву файлу");
                    SavesManager.SaveToJsonTable(table, path);
                }
                System.Environment.Exit(0);
            }
        }

        /*private override void OnClosing()
        {
            bool answer = await DisplayAlert("Підтвердження", "Ви дійсно хочете вийти?", "Так", "Ні");
            if (answer)
            {
                bool answer1 = await DisplayAlert("Збереження", "Зберігати файл?", "Так", "Ні");
                if (answer1)
                {
                    string path = await DisplayPromptAsync("Збереження файлу", "Вкажіть шлях та назву файлу");
                    SavesManager.SaveToJsonTable(table, path);
                }
                System.Environment.Exit(0);
            }
        }*/


        private async void HelpButton_Clicked(object sender, EventArgs e)
        {
            await DisplayAlert("Довідка", "Лабораторна робота 1. Студента Минька Вадима, група К-24. Варіант 11", "OK");
        }
        private void DeleteRowButton_Clicked(object sender, EventArgs e)
        {
            if (grid.RowDefinitions.Count > 1)
            {
                int lastRowIndex = grid.RowDefinitions.Count - 1;
                grid.RowDefinitions.RemoveAt(lastRowIndex);
                grid.Children.RemoveAt(lastRowIndex * (table.CurrentCountColumn + 1)); // Remove label
                for (int col = 0; col < table.CurrentCountColumn; col++)
                {
                    grid.Children.RemoveAt((lastRowIndex * table.CurrentCountColumn) + col + 1); // Remove entry
                }
            }
            table.CurrentCountRow--;
        }
        private void DeleteColumnButton_Clicked(object sender, EventArgs e)
        {
            if (grid.ColumnDefinitions.Count > 1)
            {
                int lastColumnIndex = grid.ColumnDefinitions.Count - 1;
                grid.ColumnDefinitions.RemoveAt(lastColumnIndex);
                grid.Children.RemoveAt(lastColumnIndex); // Remove label
                for (int row = 0; row < table.CurrentCountRow; row++)
                {
                    grid.Children.RemoveAt(row * (table.CurrentCountColumn + 1) + lastColumnIndex + 1); // Remove entry
                }
            }
            table.CurrentCountColumn--;
        }
        private void AddRowButton_Clicked(object sender, EventArgs e)
        {
            int newRow = grid.RowDefinitions.Count;
            // Add a new row definition
            grid.RowDefinitions.Add(new RowDefinition());
            // Add label for the row number
            var label = new Label
            {
                Text = newRow.ToString(),
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center
            };
            Grid.SetRow(label, newRow);
            Grid.SetColumn(label, 0);
            grid.Children.Add(label);
            // Add entry cells for the new row
            for (int col = 0; col < table.CurrentCountColumn; col++)
            {
                var entry = new Entry
                {
                    Text = "",
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center,
                    IsReadOnly = true,
                    WidthRequest = 100
                };
                entry.Focused += Entry_Focused;
                string cellname = GetColumnName(col + 1)+(newRow).ToString();
                var cell = new Cell(newRow, col + 1, entry, cellname);
                Calculator.GlobalScope[cellname] = 0;
                Grid.SetRow(entry, newRow);
                Grid.SetColumn(entry, col + 1);
                grid.Children.Add(entry);
            }
            table.CurrentCountRow++;
        }
        private void AddColumnButton_Clicked(object sender, EventArgs e)
        {
            int newColumn = grid.ColumnDefinitions.Count;
            // Add a new column definition
            grid.ColumnDefinitions.Add(new ColumnDefinition{ Width = 100 });
            // Add label for the column name
            var label = new Label
            {
                Text = GetColumnName(newColumn),
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center
            };
            Grid.SetRow(label, 0);
            Grid.SetColumn(label, newColumn);
            grid.Children.Add(label);
            // Add entry cells for the new column
            for (int row = 0; row < table.CurrentCountRow; row++)
            {
                var entry = new Entry
                {
                    Text = "",
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center,
                    IsReadOnly = true,
                    WidthRequest = 100
                };
                entry.Focused += Entry_Focused;
                string cellname = GetColumnName(newColumn)+(row + 1).ToString();
                var cell = new Cell(row + 1, newColumn, entry, cellname);
                Calculator.GlobalScope[cellname] = 0;
                Grid.SetRow(entry, row + 1);
                Grid.SetColumn(entry, newColumn);
                grid.Children.Add(entry);
            }
            table.CurrentCountColumn++;
        }
    }
}