using KPLN_Tools.ExecutableCommand;
using KPLN_Tools.Forms.Models.Core;
using System;
using System.ComponentModel;
using System.Data;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools.Forms.Models
{
    public sealed class ScheduleFormVM : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private DataView _dataView;
        private string _title;

        public ScheduleFormVM(Window mainWindow, ScheduleFormM model, string viewName)
        {
            MainSchIncWindow = mainWindow;
            SFModel = model ?? throw new ArgumentNullException(nameof(model));

            Title = $"KPLN_Сетка спецификации: {viewName}";
            BuildDataView();

            SetToRevitCmd = new RelayCommand<object>(SetToRevit);
            CloseWindowCmd = new RelayCommand<object>(CloseWindow);
        }

        public Window MainSchIncWindow { get; }

        /// <summary>
        /// Ссылка на модель
        /// </summary>
        public ScheduleFormM SFModel { get; }

        /// <summary>
        /// Загаловак акна / табліцы
        /// </summary>
        public string Title
        {
            get => _title;
            set
            {
                if (_title == value)
                    return;

                _title = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Дадзеныя для DataGrid
        /// </summary>
        public DataView DataView
        {
            get => _dataView;
            private set
            {
                if (_dataView == value)
                    return;

                _dataView = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Команда: Перенести в ревит
        /// </summary>
        public ICommand SetToRevitCmd { get; }

        /// <summary>
        /// Команда: Закрыть окно
        /// </summary>
        public ICommand CloseWindowCmd { get; }

        public void SetToRevit(object windObj)
        {
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ExcCmd_SchIncSetter(this));
            CloseWindow(windObj);
        }

        public void CloseWindow(object windObj)
        {
            if (windObj is Window window)
                window.Close();
        }

        /// <summary>
        /// Обновляет значение ячейки в DataView и в модели
        /// </summary>
        public void UpdateCell(DataRowView rowView, int columnIndex, ScheduleSubFormVM ssfVM)
        {
            if (rowView == null)
                return;

            // 1) Абнаўляем DataTable праз DataRowView
            var row = rowView.Row;
            if (columnIndex < 0 || columnIndex >= row.Table.Columns.Count)
                return;

            row[columnIndex] = ssfVM.FullCellData ?? string.Empty;

            // 2) Сінхранізуем з мадэллю _model.Rows
            int modelRowIndex = row.Table.Rows.IndexOf(row);
            if (modelRowIndex < 0 || modelRowIndex >= SFModel.Rows.Count)
                return;

            var modelRow = SFModel.Rows[modelRowIndex];

            // на ўсякі выпадак гарантуем памер спісу
            while (modelRow.Count <= columnIndex)
                modelRow.Add(new CellData(string.Empty, false));

            modelRow[columnIndex] = new CellData(ssfVM.FullCellData, true, true) ?? new CellData(string.Empty, false, true);
        }

        /// <summary>
        /// Инкремент значений относительно выбранной ячейки вниз по столбцу
        /// </summary>
        public void IncrementColumnBelow(DataRowView startRowView, int columnIndex, ScheduleSubFormVM ssfVM)
        {
            if (startRowView == null)
                return;

            DataRow startRow = startRowView.Row;
            DataTable table = startRow.Table;

            if (columnIndex < 0 || columnIndex >= table.Columns.Count)
                return;


            if (!int.TryParse(ssfVM.StartNumber, out int startValue))
            {
                MessageBox.Show(MainSchIncWindow, "В поле для стартовой цифры ввели НЕ цифру. Автоинкремента остановлена", "Ошибка автоинкременты", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }


            int startIndex = table.Rows.IndexOf(startRow);
            if (startIndex < 0)
                return;

            // Опускаемся вниз по столбцу
            for (int r = startIndex + 1; r < table.Rows.Count; r++)
            {
                if (!SFModel.Rows[r][columnIndex].CD_IsEditableCell)
                {
                    startIndex++;
                    continue;
                }

                int value = startValue + (r - startIndex);
                string valueWithPrefix = $"{ssfVM.Prefix}{value}";

                table.Rows[r][columnIndex] = valueWithPrefix;

                var modelRow = SFModel.Rows[r];

                // на ўсякі выпадак гарантуем памер спісу
                while (modelRow.Count <= columnIndex)
                    modelRow.Add(new CellData(string.Empty, false));
                modelRow[columnIndex] = new CellData(valueWithPrefix, true, true) ?? new CellData(string.Empty, false, true);
            }
        }

        private void BuildDataView()
        {
            DataTable table = new DataTable("Schedule");

            // Стварэнне слупкоў
            foreach (var header in SFModel.ColumnHeaders)
            {
                // калі назвы паўтараюцца – можна дадаць індэкс, каб не было дублікату
                string columnName = string.IsNullOrWhiteSpace(header.CD_Data) ? "Column" : header.CD_Data;
                if (table.Columns.Contains(columnName))
                {
                    int idx = 2;
                    string newName;
                    do
                    {
                        newName = $"{columnName}_{idx++}";
                    } while (table.Columns.Contains(newName));

                    columnName = newName;
                }

                table.Columns.Add(columnName, typeof(string));
            }

            // Даданне радкоў
            foreach (var row in SFModel.Rows)
            {
                var dataRow = table.NewRow();
                for (int i = 0; i < SFModel.ColumnCount; i++)
                {
                    string val = i < row.Count ? row[i].CD_Data : string.Empty;
                    dataRow[i] = val;
                }
                table.Rows.Add(dataRow);
            }

            DataView = table.DefaultView;
        }

        private void OnPropertyChanged([CallerMemberName] string propName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
    }
}