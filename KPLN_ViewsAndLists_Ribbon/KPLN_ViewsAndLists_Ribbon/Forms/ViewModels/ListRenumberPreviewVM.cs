using Autodesk.Revit.DB;
using KPLN_ViewsAndLists_Ribbon.Forms.Entities;
using KPLN_ViewsAndLists_Ribbon.Common.Lists;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_ViewsAndLists_Ribbon.Forms.MVVM
{
    public sealed class ListRenumberPreviewVM : INotifyPropertyChanged
    {
        private readonly Document _document;
        private readonly RenameMode _mode;
        private readonly string _unicodeCode;
        private readonly string _unicodeName;
        private HashSet<string> _otherNumbers;
        private string _status;
        private bool _isContinuingNumbers;
        private bool _hasErrors;
        private bool _busy;
        private string _contextError;
        public bool CanApply => !_hasErrors && !_busy && _contextError == null;
        public bool CanEdit => !_busy;

        internal void SetBusy(bool value)
        {
            _busy = value;
            OnPropertyChanged(nameof(CanApply));
            OnPropertyChanged(nameof(CanEdit));
            OnPropertyChanged(nameof(CanContinueNumbering));
        }

        internal void SetContextError(string error)
        {
            if (_contextError == error) return;
            _contextError = error;
            Validate();
        }

        // Вызывается только из контекста Revit; редактирование таблицы использует снимок.
        internal void RefreshModel()
        {
            foreach (var row in Rows)
            {
                row.ModelError = !row.Sheet.IsValidObject
                    ? "Лист удалён. Сформируйте предпросмотр заново."
                    : row.Sheet.SheetNumber != row.OriginalNumber
                        ? "Номер листа изменился. Сформируйте предпросмотр заново." : null;
                if (row.ModelError == null && HasParameter)
                {
                    var parameter = row.Sheet.LookupParameter(ParameterName);
                    row.TargetParameter = parameter;
                    row.CanWriteParameter = parameter != null && !parameter.IsReadOnly && parameter.StorageType == StorageType.String;
                    if (row.CanWriteParameter && (parameter.AsString() ?? string.Empty) != row.OriginalParameterValue)
                        row.ModelError = "Параметр листа изменился. Сформируйте предпросмотр заново.";
                }
            }
            if (Rows.All(row => row.Sheet.IsValidObject)) ReadOtherNumbers();
            Validate();
        }

        internal ListRenumberPreviewVM(Document document, IEnumerable<ListRenumberPreviewRow> rows,
            RenameMode mode, string parameterName, string unicodeCode, string unicodeName)
        {
            _document = document;
            _mode = mode;
            _unicodeCode = unicodeCode;
            _unicodeName = unicodeName;
            Rows = new ObservableCollection<ListRenumberPreviewRow>(rows);
            ParameterName = parameterName;
            ReadOtherNumbers();

            foreach (ListRenumberPreviewRow row in Rows)
                row.PropertyChanged += RowChanged;

            Validate();
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public ObservableCollection<ListRenumberPreviewRow> Rows { get; }
        public string ParameterName { get; }
        public bool HasParameter => ParameterName != null;
        public bool IsNumberReadOnly => _mode == RenameMode.RefreshParam;
        public bool CanContinueNumbering => !IsNumberReadOnly && !_busy;
        public string Help => IsNumberReadOnly
            ? "Редактируйте столбец «Новое значение параметра». Номера листов в этом режиме не изменяются."
            : "Редактируйте только номер. Приставки и Unicode сохраняются из предыдущего окна." + (HasParameter
                ? " При изменении номера значение параметра пересчитывается; его также можно исправить вручную."
                : string.Empty);

        public string Status
        {
            get => _status;
            private set { _status = value; OnPropertyChanged(); }
        }

        private void RowChanged(object sender, PropertyChangedEventArgs e)
        {
            if (_isContinuingNumbers) return;

            if (e.PropertyName == nameof(ListRenumberPreviewRow.NewNumber)
                || e.PropertyName == nameof(ListRenumberPreviewRow.EditableNumber)
                || e.PropertyName == nameof(ListRenumberPreviewRow.NewParameterValue))
                Validate();
        }

        internal bool ContinueNumbering(IList<ListRenumberPreviewRow> displayedRows, ListRenumberPreviewRow selectedRow)
        {
            if (IsNumberReadOnly)
            {
                Status = "В этом режиме изменяется только параметр штампа.";
                return false;
            }

            int selectedIndex = selectedRow == null ? -1 : displayedRows.IndexOf(selectedRow);
            if (selectedIndex < 0)
            {
                Status = "Выберите ячейку в столбце «Новый номер», от которой нужно продолжить нумерацию.";
                return false;
            }
            if (selectedIndex == displayedRows.Count - 1)
            {
                Status = "Ниже выбранной строки нет листов для продолжения нумерации.";
                return false;
            }

            // Сначала готовим весь результат; при неподходящем формате таблица не изменяется.
            var changes = new List<KeyValuePair<ListRenumberPreviewRow, string>>();
            ListRenumberPreviewRow currentRow = selectedRow;
            try
            {
                if (selectedRow.NumberEditError != null) throw new FormatException(selectedRow.NumberEditError);
                long next = ListRenumberSequenceNumber.Parse(selectedRow.NewNumber).Value;
                for (int index = selectedIndex + 1; index < displayedRows.Count; index++)
                {
                    currentRow = displayedRows[index];
                    if (currentRow.NumberEditError != null) throw new FormatException(currentRow.NumberEditError);
                    next = checked(next + 1);
                    string number = ListRenumberSequenceNumber.Parse(currentRow.NewNumber).WithValue(next);
                    changes.Add(new KeyValuePair<ListRenumberPreviewRow, string>(currentRow, number));
                }
            }
            catch (Exception ex) when (ex is FormatException || ex is OverflowException)
            {
                Status = $"Нумерация не изменена. Лист «{currentRow.SheetName}» ({currentRow.OriginalNumber}): "
                    + (ex is OverflowException ? "слишком большое значение номера." : ex.Message);
                return false;
            }

            _isContinuingNumbers = true;
            try
            {
                foreach (var change in changes)
                    change.Key.NewNumber = change.Value;
            }
            finally
            {
                _isContinuingNumbers = false;
                Validate();
            }
            return true;
        }

        private void ReadOtherNumbers()
        {
            var selectedIds = new HashSet<ElementId>(Rows.Select(row => row.Sheet.Id));
            _otherNumbers = new HashSet<string>(new FilteredElementCollector(_document)
                .OfClass(typeof(ViewSheet)).Cast<ViewSheet>()
                .Where(sheet => !selectedIds.Contains(sheet.Id)).Select(sheet => sheet.SheetNumber), StringComparer.Ordinal);
        }

        private bool Validate()
        {
            var duplicates = new HashSet<string>(Rows.GroupBy(row => row.NewNumber, StringComparer.Ordinal)
                .Where(group => group.Count() > 1).Select(group => group.Key), StringComparer.Ordinal);

            foreach (ListRenumberPreviewRow row in Rows)
            {
                string error = row.NumberEditError ?? row.CalculationError ?? row.ParameterCalculationError;
                if (string.IsNullOrWhiteSpace(row.NewNumber))
                    error = "Укажите номер листа.";
                else if (duplicates.Contains(row.NewNumber))
                    error = "Номер повторяется в таблице.";
                else if (_otherNumbers.Contains(row.NewNumber))
                    error = "Номер занят другим листом проекта.";

                if (HasParameter && !row.CanWriteParameter)
                    error = "Параметр отсутствует, доступен только для чтения или не является текстовым.";

                row.Error = row.ModelError ?? error;
            }

            int errors = Rows.Count(row => !string.IsNullOrEmpty(row.Error));
            _hasErrors = errors != 0;
            OnPropertyChanged(nameof(CanApply));
            Status = _contextError ?? (errors == 0 ? $"Листов: {Rows.Count}. Можно применить изменения."
                : $"Листов: {Rows.Count}. Исправьте строки с ошибками: {errors}.");
            return errors == 0;
        }

        internal bool TryApply()
        {
            try
            {
                RefreshModel();
                if (!Validate() || _contextError != null) return false;

                if (Rows.Any(row => !row.Sheet.IsValidObject || row.Sheet.SheetNumber != row.OriginalNumber))
                {
                    Status = "Листы изменились после открытия окна. Закройте предпросмотр и сформируйте его заново.";
                    return false;
                }

                using (Transaction transaction = new Transaction(_document, $"KPLN: {ListRenumberVM.PluginName}"))
                {
                    transaction.Start();

                    // Освобождаем исходные номера, чтобы можно было поменять номера листов местами.
                    var changed = Rows.Where(row => row.NewNumber != row.OriginalNumber).ToArray();
                    string temporaryPrefix = "KPLN_" + Guid.NewGuid().ToString("N") + "_";
                    for (int index = 0; index < changed.Length; index++)
                        changed[index].Sheet.SheetNumber = temporaryPrefix + index;

                    foreach (ListRenumberPreviewRow row in changed)
                        row.Sheet.SheetNumber = row.NewNumber;

                    // Записываем именно те значения, которые пользователь проверил в таблице.
                    foreach (ListRenumberPreviewRow row in Rows)
                    {
                        if (HasParameter && row.TargetParameter.AsString() != (row.NewParameterValue ?? string.Empty))
                        {
                            if (!row.TargetParameter.Set(row.NewParameterValue ?? string.Empty))
                                throw new InvalidOperationException($"Не удалось заполнить параметр листа {row.NewNumber}.");
                        }

                        if (_mode == RenameMode.Unicode || _mode == RenameMode.ClearedNumber)
                        {
                            Parameter unicodeParameter = row.Sheet.get_Parameter(new Guid("09b934d4-81b3-4aff-b37e-d20dfbc1ac8e"));
                            if (unicodeParameter != null)
                            {
                                int count = _mode == RenameMode.Unicode
                                    ? row.NewNumber.TakeWhile(character => character == _unicodeCode[0]).Count() : 0;
                                string value = count == 0 ? string.Empty : count == 1 ? _unicodeName : $"{count}X{_unicodeName}";
                                if (unicodeParameter.AsString() != value && !unicodeParameter.Set(value))
                                    throw new InvalidOperationException($"Не удалось обновить обозначение Unicode листа {row.NewNumber}.");
                            }
                        }
                    }

                    if (transaction.Commit() != TransactionStatus.Committed)
                    {
                        Status = "Revit не подтвердил запись изменений. Проверьте сообщения Revit.";
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Status = "Изменения не применены: " + ex.Message;
                return false;
            }
        }

        private void OnPropertyChanged([CallerMemberName] string name = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
