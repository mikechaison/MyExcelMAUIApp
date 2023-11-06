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
        public Table table { get; set; } //таблиця
        public Cell currentcell { get; set; } = null; //обрана клітина
        IFileSaver fileSaver { get; }
        CancellationTokenSource cancellationTokenSource { get; } = new CancellationTokenSource();

        public MainPage(IFileSaver fileSaver)
        {
            table = new Table(CountRow, CountColumn);
            this.fileSaver = fileSaver;
            InitializeComponent();
            CreateGrid();
        }

        //СІтворення таблиці
        private void CreateGrid()
        {
            AddColumnsAndColumnLabels();
            AddRowsAndCellEntries();
        }

        //Зміна таблиці
        private void ChangeGrid()
        {
            var nowRows = grid.RowDefinitions.Count;
            var nowColumns = grid.ColumnDefinitions.Count - 1;
            var nextRows = table.CurrentCountRow;
            var nextColumns = table.CurrentCountColumn;
            if (nextRows >= nowRows)
            {
                for (int i = 0; i < nextRows - nowRows; i++)
                {
                    AddRow(false);
                }
            }
            else
            {
                for (int i = 0; i < nowRows - nextRows; i++)
                {
                    DeleteRow(false);
                }
            }

            if (nextColumns >= nowColumns)
            {
                for (int i = 0; i < nextColumns - nowColumns; i++)
                {
                    AddColumn(false);
                }
            }
            else
            {
                for (int i = 0; i < nowColumns - nextColumns; i++)
                {
                    DeleteColumn(false);
                }
            }
        }

        //Оновлення таблиці
        private void UpdateGrid()
        {
            ChangeGrid();
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

        //Заповнює таблицю після зчитування файлу
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
        
        //Перераховує всю таблицю
        private void CalculateTable()
        {
            foreach (Cell cell in table.cells)
            {
                cell.Calculate();
            }
        }

        //Отримує ім'я колонки
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

        //Викликається, коли користувач заходить у клітину (набуває фокус)
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

        //Викликається, коли користувач виходить з рядка для введення тексту
        private void TextInput_Return(object sender, FocusEventArgs e)
        {
            if (currentcell != null)
            {
                var calc = table.TryCalculate(currentcell, textInput.Text);
                if (!calc)
                {
                    textInput.Text = "";
                    CycleError();
                    CalculateTable();
                }
                else
                {
                    try
                    {
                        currentcell.Calculate();
                        table.RecalculateRecursively(currentcell);
                    }
                    catch (ArgumentException argex)
                    {
                        textInput.Text = "";
                        table.ClearCellData(currentcell);
                        ExpressionError(argex.Message);
                        CalculateTable();
                    }
                }
            }
        }

        private async void CycleError()
        {
            await DisplayAlert("Помилка", "Виявлена циклічна залежність!", "OK");
        }

        private async void ExpressionError(string text)
        {
            await DisplayAlert("Помилка", text, "OK");
        }

        private async void SaveButton_Clicked(object sender, EventArgs e)
        {
            using var stream = new MemoryStream(Encoding.Default.GetBytes("Text"));
            var path = await fileSaver.SaveAsync("table.json", stream, cancellationTokenSource.Token);
            if (path != null)
            {
                SavesManager.SaveToJsonTable(table, path.FilePath);
            }
        }
        
        private async void ReadButton_Clicked(object sender, EventArgs e)
        {
            var customFileType = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
            {
                { DevicePlatform.WinUI, new[] { ".json" } }
            });
            var result = await FilePicker.PickAsync(new PickOptions
            {
                PickerTitle = "Оберіть файл",
                FileTypes = customFileType
            });
            if (result != null)
            {
                table = SavesManager.ReadJsonTable(result.FullPath);
                UpdateGrid();
            }
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
        private async void HelpButton_Clicked(object sender, EventArgs e)
        {
            await DisplayAlert("Довідка", "Лабораторна робота 1. Студента Минька Вадима, група К-24. Варіант 11", "OK");
        }
        private void DeleteRowButton_Clicked(object sender, EventArgs e)
        {
            DeleteRow(true); //видалення зі зміною кількості рядків
        }
        private void DeleteRow(bool add)
        {
            if (grid.RowDefinitions.Count > 1)
            {
                int lastRowIndex = grid.RowDefinitions.Count - 1;
                grid.RowDefinitions.RemoveAt(lastRowIndex);
                var ce = grid.Children.ElementAt(lastRowIndex * (table.CurrentCountColumn + 1) + table.CurrentCountColumn);
                grid.Children.RemoveAt(lastRowIndex * (table.CurrentCountColumn + 1) + table.CurrentCountColumn); // Remove label
                for (int col = 0; col < table.CurrentCountColumn; col++)
                {
                    grid.Children.RemoveAt(lastRowIndex * (table.CurrentCountColumn + 1) + table.CurrentCountColumn); // Remove entry
                    Cell c = table.FindCellByEntry(table.CurrentCountRow, col + 1);
                    table.cells.Remove(c);
                }
            }
            if (add)
            {
                table.CurrentCountRow--;
            }
        }

        private void DeleteColumnButton_Clicked(object sender, EventArgs e)
        {
            DeleteColumn(true); //видалення зі зміною кількості колонок
        }

        private void DeleteColumn(bool add)
        {
            if (grid.ColumnDefinitions.Count > 1)
            {
                int lastColumnIndex = grid.ColumnDefinitions.Count - 1;
                grid.ColumnDefinitions.RemoveAt(lastColumnIndex);
                for (int row = grid.RowDefinitions.Count - 1; row >= 0 ; row--)
                {
                    grid.Children.RemoveAt(row * (table.CurrentCountColumn + 1) + 2 * table.CurrentCountColumn); // Remove entry
                    Cell c = table.FindCellByEntry(row + 1, table.CurrentCountColumn);
                    table.cells.Remove(c);
                }
                grid.Children.RemoveAt(lastColumnIndex - 1);
            }
            if (add)
            {
                table.CurrentCountColumn--;
            }
        }

        private void AddRow(bool add_cell)
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
                if (add_cell) //якщо це не при зчитуванні файлу
                {
                    string cellname = GetColumnName(col + 1)+(newRow).ToString();
                    var cell = new Cell(newRow, col + 1, entry, cellname);
                    Calculator.GlobalScope[cellname] = 0;
                    table.AddCellToTable(cell);
                }
                Grid.SetRow(entry, newRow);
                Grid.SetColumn(entry, col + 1);
                grid.Children.Add(entry);
            }
            if (add_cell)
            {
                table.CurrentCountRow++;
            }
        }

        private void AddRowButton_Clicked(object sender, EventArgs e)
        {
            AddRow(true); //додавання зі зміною кількості рядків
        }

        private void AddColumn(bool add_cell)
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
                if (add_cell) //якщо це не при зчитуванні файлу
                {
                    string cellname = GetColumnName(newColumn)+(row + 1).ToString();
                    var cell = new Cell(row + 1, newColumn, entry, cellname);
                    Calculator.GlobalScope[cellname] = 0;
                    table.AddCellToTable(cell);
                    
                }
                Grid.SetRow(entry, row + 1);
                Grid.SetColumn(entry, newColumn);
                grid.Children.Add(entry);
            }
            if (add_cell)
            {
                table.CurrentCountColumn++;
            }
        }

        private void AddColumnButton_Clicked(object sender, EventArgs e)
        {
            AddColumn(true); //додавання зі зміною кількості стовпців
        }
        
    }
}