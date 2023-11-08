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
        public Table OurTable { get; set; } //таблиця
        public Cell CurrentCell { get; set; } = null; //обрана клітина
        public Dictionary<int, Label> GridRowLabels { get; set; } = new Dictionary<int, Label>(); //лейбли для рядків
        public Dictionary<int, Label> GridColumnLabels { get; set; } = new Dictionary<int, Label>(); //лейбли для стовпчиків
        public Dictionary<string, Entry> GridEntries { get; set; } = new Dictionary<string, Entry>(); //клітинки за іменем
        IFileSaver fileSaver { get; }
        CancellationTokenSource cancellationTokenSource { get; } = new CancellationTokenSource();

        public MainPage(IFileSaver fileSaver)
        {
            OurTable = new Table(CountRow, CountColumn);
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
            var NowRows = grid.RowDefinitions.Count;
            var NowColumns = grid.ColumnDefinitions.Count - 1;
            var NextRows = OurTable.CurrentCountRow;
            var NextColumns = OurTable.CurrentCountColumn;
            if (NextRows >= NowRows)
            {
                for (int i = 0; i < NextRows - NowRows; i++)
                {
                    AddRow(false);
                }
            }
            else
            {
                for (int i = 0; i < NowRows - NextRows; i++)
                {
                    DeleteRow(false);
                }
            }

            if (NextColumns >= NowColumns)
            {
                for (int i = 0; i < NextColumns - NowColumns; i++)
                {
                    AddColumn(false);
                }
            }
            else
            {
                for (int i = 0; i < NowColumns - NextColumns; i++)
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
            for (int Col = 0; Col < OurTable.CurrentCountColumn + 1; Col++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition{ Width = 100 });
                if (Col > 0)
                {
                    var OurLabel = new Label
                    {
                        Text = GetColumnName(Col),
                        VerticalOptions = LayoutOptions.Center,
                        HorizontalOptions = LayoutOptions.Center
                    };
                    Grid.SetRow(OurLabel, 0);
                    Grid.SetColumn(OurLabel, Col);
                    grid.Children.Add(OurLabel);
                    GridColumnLabels[Col] = OurLabel;
                }
                else
                {
                    grid.ColumnDefinitions[Col].Width = 40;
                }
            }
        }
        private void AddRowsAndCellEntries()
        {
            // Додати рядки, підписи для рядків та комірки
            for (int Row = 0; Row < OurTable.CurrentCountRow; Row++)
            {
                grid.RowDefinitions.Add(new RowDefinition());
                // Додати підпис для номера рядка
                var OurLabel = new Label
                {
                    Text = (Row + 1).ToString(),
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center
                };
                Grid.SetRow(OurLabel, Row + 1);
                Grid.SetColumn(OurLabel, 0);
                grid.Children.Add(OurLabel);
                GridRowLabels[Row + 1] = OurLabel;
                // Додати комірки (Entry) для вмісту
                for (int Col = 0; Col < OurTable.CurrentCountColumn; Col++)
                {
                    var OurEntry = new Entry
                    {
                        Text = "",
                        VerticalOptions = LayoutOptions.Center,
                        HorizontalOptions = LayoutOptions.Center,
                        IsReadOnly = true,
                        WidthRequest = 100
                    };
                    OurEntry.Focused += Entry_Focused; // обробник події Focused
                    string CellName = GetColumnName(Col + 1)+(Row + 1).ToString();
                    var OurCell = new Cell(Row + 1, Col + 1, OurEntry, CellName);
                    Calculator.GlobalScope[CellName] = 0;
                    OurTable.AddCellToTable(OurCell);
                    Grid.SetRow(OurEntry, Row + 1);
                    Grid.SetColumn(OurEntry, Col + 1);
                    grid.Children.Add(OurEntry);
                    GridEntries[CellName] = OurEntry;
                    OurTable.CellNames[CellName] = OurCell;
                }
            }
        }

        //Заповнює таблицю після зчитування файлу
        private void FillTable()
        {
            foreach (Cell OurCell in OurTable.Cells)
            {
                Entry CellEntry = GridEntries[OurCell.CellName];
                CellEntry.Text = OurCell.Expression;
                OurCell.CellEntry = CellEntry;
            }
        }
        
        //Перераховує всю таблицю
        private void CalculateTable()
        {
            foreach (Cell OurCell in OurTable.Cells)
            {
                OurCell.Calculate();
            }
        }

        //Отримує ім'я колонки
        private string GetColumnName(int ColIndex)
        {
            int Dividend = ColIndex;
            string ColumnName = string.Empty;
            while (Dividend > 0)
            {
                int Modulo = (Dividend - 1) % 26;
                ColumnName = Convert.ToChar(65 + Modulo) + ColumnName;
                Dividend = (Dividend - Modulo) / 26;
            }
            return ColumnName;
        }

        //Викликається, коли користувач заходить у клітину (набуває фокус)
        private void Entry_Focused(object sender, FocusEventArgs e)
        {
            if (CurrentCell != null)
            {
                CurrentCell.CellEntry.BackgroundColor = Colors.White;
            }
            var OurEntry = (Entry)sender;
            var Row = Grid.GetRow(OurEntry) - 1;
            var Col = Grid.GetColumn(OurEntry) - 1;
            var CellName = GetColumnName(Col + 1)+(Row + 1).ToString();
            CurrentCell = OurTable.CellNames[CellName];
            TextInput.Text = CurrentCell.Expression;
            CellLabel.Text = CurrentCell.CellName;
            OurEntry.BackgroundColor = Colors.LightYellow;
        }

        //Викликається, коли користувач виходить з рядка для введення тексту
        private void TextInput_Return(object sender, FocusEventArgs e)
        {
            if (CurrentCell != null)
            {
                var calc = OurTable.TryCalculate(CurrentCell, TextInput.Text);
                if (!calc)
                {
                    TextInput.Text = "";
                    CycleError();
                    CalculateTable();
                }
                else
                {
                    try
                    {
                        CurrentCell.Calculate();
                        OurTable.RecalculateRecursively(CurrentCell);
                    }
                    catch (ArgumentException argex)
                    {
                        TextInput.Text = "";
                        OurTable.ClearCellData(CurrentCell);
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

        private async void ExpressionError(string Text)
        {
            await DisplayAlert("Помилка", Text, "OK");
        }

        private async void SaveFile()
        {
            using var stream = new MemoryStream(Encoding.Default.GetBytes("Text"));
            var path = await fileSaver.SaveAsync("OurTable.json", stream, cancellationTokenSource.Token);
            if (path.FilePath != null)
            {
                SavesManager.SaveToJsonTable(OurTable, path.FilePath);
            }
        }

        private async void SaveButton_Clicked(object sender, EventArgs e)
        {
            SaveFile();
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
                OurTable = SavesManager.ReadJsonTable(result.FullPath);
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
                    SaveFile();
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

        private void DeleteRow(bool Add)
        {
            if (grid.RowDefinitions.Count > 1)
            {
                int LastRowIndex = grid.RowDefinitions.Count - 1;
                grid.RowDefinitions.RemoveAt(LastRowIndex);
                grid.Children.Remove(GridRowLabels[LastRowIndex]); // Remove label
                for (int Col = 0; Col < OurTable.CurrentCountColumn; Col++) //Remove entries
                {
                    string CellName = GetColumnName(Col + 1)+LastRowIndex.ToString();
                    Entry CellEntry = GridEntries[CellName];
                    grid.Children.Remove(CellEntry);
                    if (Add)
                    {
                        Cell OurCell = OurTable.CellNames[CellName];
                        OurTable.Cells.Remove(OurCell);
                    }
                }
            }
            if (Add)
            {
                OurTable.CurrentCountRow--;
            }
        }

        private void DeleteColumnButton_Clicked(object sender, EventArgs e)
        {
            DeleteColumn(true); //видалення зі зміною кількості колонок
        }

        private void DeleteColumn(bool Add)
        {
            if (grid.ColumnDefinitions.Count > 1)
            {
                int LastColumnIndex = grid.ColumnDefinitions.Count - 1;
                grid.ColumnDefinitions.RemoveAt(LastColumnIndex);
                grid.Children.Remove(GridColumnLabels[LastColumnIndex]); // Remove label
                for (int Row = OurTable.CurrentCountRow; Row > 0 ; Row--) //Remove entries
                {
                    string CellName = GetColumnName(LastColumnIndex)+Row.ToString();
                    Entry CellEntry = GridEntries[CellName];
                    grid.Children.Remove(CellEntry);
                    if (Add)
                    {
                        Cell OurCell = OurTable.CellNames[CellName];
                        OurTable.Cells.Remove(OurCell);
                    }
                }
            }
            if (Add)
            {
                OurTable.CurrentCountColumn--;
            }
        }

        private void AddRow(bool AddCell)
        {
            int NewRow = grid.RowDefinitions.Count;
            // Add a new Row definition
            grid.RowDefinitions.Add(new RowDefinition());
            // Add laberl for the Row number
            var OurLabel = new Label
            {
                Text = NewRow.ToString(),
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center
            };
            Grid.SetRow(OurLabel, NewRow);
            Grid.SetColumn(OurLabel, 0);
            grid.Children.Add(OurLabel);
            GridRowLabels[NewRow] = OurLabel;
            // Add OurEntry cells for the new Row
            for (int Col = 0; Col < OurTable.CurrentCountColumn; Col++)
            {
                var OurEntry = new Entry
                {
                    Text = "",
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center,
                    IsReadOnly = true,
                    WidthRequest = 100
                };
                OurEntry.Focused += Entry_Focused;
                string CellName = GetColumnName(Col + 1)+NewRow.ToString();
                GridEntries[CellName] = OurEntry;
                if (AddCell) //якщо це не при зчитуванні файлу
                {
                    var OurCell = new Cell(NewRow, Col + 1, OurEntry, CellName);
                    Calculator.GlobalScope[CellName] = 0;
                    OurTable.AddCellToTable(OurCell);
                    OurTable.CellNames[CellName] = OurCell;
                }
                Grid.SetRow(OurEntry, NewRow);
                Grid.SetColumn(OurEntry, Col + 1);
                grid.Children.Add(OurEntry);
                
            }
            if (AddCell)
            {
                OurTable.CurrentCountRow++;
            }
        }

        private void AddRowButton_Clicked(object sender, EventArgs e)
        {
            AddRow(true); //додавання зі зміною кількості рядків
        }

        private void AddColumn(bool AddCell)
        {
            int NewColumn = grid.ColumnDefinitions.Count;
            // Add a new column definition
            grid.ColumnDefinitions.Add(new ColumnDefinition{ Width = 100 });
            // Add label for the column name
            var OurLabel = new Label
            {
                Text = GetColumnName(NewColumn),
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center
            };
            Grid.SetRow(OurLabel, 0);
            Grid.SetColumn(OurLabel, NewColumn);
            grid.Children.Add(OurLabel);
            GridColumnLabels[NewColumn] = OurLabel;
            // Add OurEntry cells for the new column
            for (int Row = 0; Row < OurTable.CurrentCountRow; Row++)
            {
                var OurEntry = new Entry
                {
                    Text = "",
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center,
                    IsReadOnly = true,
                    WidthRequest = 100
                };
                OurEntry.Focused += Entry_Focused;
                string CellName = GetColumnName(NewColumn)+(Row + 1).ToString();
                GridEntries[CellName] = OurEntry;
                if (AddCell) //якщо це не при зчитуванні файлу
                {
                    var OurCell = new Cell(Row + 1, NewColumn, OurEntry, CellName);
                    Calculator.GlobalScope[CellName] = 0;
                    OurTable.AddCellToTable(OurCell);
                    OurTable.CellNames[CellName] = OurCell;
                }
                Grid.SetRow(OurEntry, Row + 1);
                Grid.SetColumn(OurEntry, NewColumn);
                grid.Children.Add(OurEntry);
            }
            if (AddCell)
            {
                OurTable.CurrentCountColumn++;
            }
        }

        private void AddColumnButton_Clicked(object sender, EventArgs e)
        {
            AddColumn(true); //додавання зі зміною кількості стовпців
        }
        
    }
}