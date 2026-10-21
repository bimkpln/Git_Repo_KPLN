using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ViewsAndLists_Ribbon.Common.Lists;
using KPLN_ViewsAndLists_Ribbon.Forms.Commands;
using KPLN_ViewsAndLists_Ribbon.Forms.Entities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace KPLN_ViewsAndLists_Ribbon.Forms.MVVM
{
    // Имена параметров читаются в контексте Revit, WPF получает готовый снимок.
    public sealed class ListRenumberParameter
    {
        internal Parameter Parameter { get; }
        public string Name { get; }
        public ListRenumberParameter(Parameter parameter)
        {
            Parameter = parameter;
            Name = parameter.Definition.Name;
        }
        public static implicit operator ListRenumberParameter(Parameter parameter)
            => parameter == null ? null : new ListRenumberParameter(parameter);
    }

    public enum RenameMode
    {
        Unicode,
        Prefix,
        OnlyRenumber,
        RefreshParam,
        ClearedNumber
    }

    public sealed class ListRenumberVM : INotifyPropertyChanged
    {
        /// <summary>
        /// Имя транзакции для анализа на наличие
        /// </summary>
        public static readonly string PluginName = "Перенумеровать листы";

        public event PropertyChangedEventHandler PropertyChanged;

        private readonly UIApplication _uiapp;
        private readonly ViewSheet[] _shetsToRenumber;

        private RenameMode _selectedMode;
        private bool _isUnicode = false;
        private bool _isPrefix = false;
        private bool _prefix_IsParamUpdate;
        private bool _isOnlyRenumber = false;
        private bool _isRefreshParam = false;
        private bool _isClearRenumber = false;

        private ListRenumberParameter _prefix_SelectedTitleBlockParameter;
        private ListRenumberParameter _onlyRenumber_SelectedTitleBlockParameter;
        private ListRenumberParameter _refreshParam_SelectedTitleBlockParameter;

        private UniEntity _unicode_SelectedUnicode;
        private bool _unicode_IsChangePrefixToUnicode;
        private bool _unicode_IsSetNumberOfUnicodes;
        private string _unicdoe_CountOfUnicodes = "1";

        private string _prefix_Text;

        private string _onlyRenumber_StartNumber = "0";
        private bool _onlyRenumber_IsParamUpdate;

        public ListRenumberVM(UIApplication uiapp, IEnumerable<ViewSheet> shetsToRenumber, IEnumerable<Parameter> tBlockParams)
        {
            _uiapp = uiapp;
            _shetsToRenumber = shetsToRenumber.ToArray();

            Unicode_UnicodeList = new ObservableCollection<UniEntity>()
            {
                new UniEntity()
                {
                    Name = UniDecCodes.LRE.ToString(),
                    Code = "‪",
                    DecCode = (int)UniDecCodes.LRE
                },
                new UniEntity()
                {
                    Name = UniDecCodes.LRO.ToString(),
                    Code = "‭",
                    DecCode = (int)UniDecCodes.LRO
                },
                new UniEntity()
                {
                    Name = UniDecCodes.PDF.ToString(),
                    Code = "‬",
                    DecCode = (int)UniDecCodes.PDF
                },
                new UniEntity()
                {
                    Name = UniDecCodes.RS.ToString(),
                    Code = "",
                    DecCode = (int)UniDecCodes.RS
                },
                new UniEntity()
                {
                    Name = UniDecCodes.US.ToString(),
                    Code = "",
                    DecCode = (int)UniDecCodes.US
                }
            };
            TitleBlockParameters = new ObservableCollection<ListRenumberParameter>(tBlockParams.Select(p => new ListRenumberParameter(p)));

            OkCommand = new RelayCommand<object>(_ =>
            {
                if (RequestApiAction != null) RequestApiAction(OnPreview);
                else OnPreview();
            });
        }

        public RenameMode SelectedMode
        {
            get => _selectedMode;
            set
            {
                if (_selectedMode != value)
                {
                    _selectedMode = value;
                    OnPropertyChanged();
                }
            }
        }

        // 0. Юникоды
        public bool IsUnicode
        {
            get => _isUnicode;
            set
            {
                _isUnicode = value;
                OnPropertyChanged();

                UpdateSelectedMode();
            }
        }

        public ObservableCollection<UniEntity> Unicode_UnicodeList { get; set; }

        public ObservableCollection<ListRenumberParameter> TitleBlockParameters { get; set; }

        public UniEntity Unicode_SelectedUnicode
        {
            get => _unicode_SelectedUnicode;
            set => Set(ref _unicode_SelectedUnicode, value);
        }

        public bool Unicode_IsChangePrefixToUnicode
        {
            get => _unicode_IsChangePrefixToUnicode;
            set => Set(ref _unicode_IsChangePrefixToUnicode, value);
        }

        public bool Unicode_IsSetNumberOfUnicodes
        {
            get => _unicode_IsSetNumberOfUnicodes;
            set => Set(ref _unicode_IsSetNumberOfUnicodes, value);
        }

        public string Unicdoe_CountOfUnicodes
        {
            get => _unicdoe_CountOfUnicodes;
            set => Set(ref _unicdoe_CountOfUnicodes, value);
        }

        // 1. Префиксы
        public bool IsPrefix
        {
            get => _isPrefix;
            set
            {
                _isPrefix = value;
                OnPropertyChanged();

                UpdateSelectedMode();
            }
        }

        public ListRenumberParameter Prefix_SelectedTitleBlockParameter
        {
            get => _prefix_SelectedTitleBlockParameter;
            set => Set(ref _prefix_SelectedTitleBlockParameter, value);
        }

        public string Prefix_Text
        {
            get => _prefix_Text;
            set => Set(ref _prefix_Text, value);
        }

        public bool Prefix_IsParamUpdate
        {
            get => _prefix_IsParamUpdate;
            set => Set(ref _prefix_IsParamUpdate, value);
        }

        // 2. Только перенумерация
        public bool IsOnlyRenumber
        {
            get => _isOnlyRenumber;
            set
            {
                _isOnlyRenumber = value;
                OnPropertyChanged();

                UpdateSelectedMode();
            }
        }

        public string OnlyRenumber_StartNumber
        {
            get => _onlyRenumber_StartNumber;
            set => Set(ref _onlyRenumber_StartNumber, value);
        }

        public bool OnlyRenumber_IsParamUpdate
        {
            get => _onlyRenumber_IsParamUpdate;
            set => Set(ref _onlyRenumber_IsParamUpdate, value);
        }

        public ListRenumberParameter OnlyRenumber_SelectedTitleBlockParameter
        {
            get => _onlyRenumber_SelectedTitleBlockParameter;
            set => Set(ref _onlyRenumber_SelectedTitleBlockParameter, value);
        }

        // 3. Обновление параметра
        public bool IsRefreshParam
        {
            get => _isRefreshParam;
            set
            {
                _isRefreshParam = value;
                OnPropertyChanged();

                UpdateSelectedMode();
            }
        }

        public ListRenumberParameter RefreshParam_SelectedTitleBlockParameter
        {
            get => _refreshParam_SelectedTitleBlockParameter;
            set => Set(ref _refreshParam_SelectedTitleBlockParameter, value);
        }

        // 4. Очистка нумерации
        public bool IsClearedNumber
        {
            get => _isClearRenumber;
            set
            {
                _isClearRenumber = value;
                OnPropertyChanged();

                UpdateSelectedMode();
            }
        }

        // 5. Управление
        public ICommand OkCommand { get; set; }

        public ICommand CancelCommand { get; set; }

        public Action CloseAction { get; set; }
        internal Action<Action> RequestApiAction { get; set; }

        public Func<ListRenumberPreviewVM, bool?> PreviewAction { get; set; }

        private void OnPreview()
        {
            UIDocument uidoc = _uiapp.ActiveUIDocument;
            if (uidoc == null) return;

            try
            {
                if (!IsUnicode && !IsPrefix && !IsOnlyRenumber && !IsRefreshParam && !IsClearedNumber)
                    throw new InvalidOperationException("Выберите режим нумерации.");

                Parameter parameter = null;
                bool updateParameter = SelectedMode == RenameMode.RefreshParam
                    || SelectedMode == RenameMode.Prefix && Prefix_IsParamUpdate
                    || SelectedMode == RenameMode.OnlyRenumber && OnlyRenumber_IsParamUpdate;
                if (updateParameter)
                {
                    parameter = (SelectedMode == RenameMode.RefreshParam ? RefreshParam_SelectedTitleBlockParameter
                        : SelectedMode == RenameMode.Prefix ? Prefix_SelectedTitleBlockParameter : OnlyRenumber_SelectedTitleBlockParameter)?.Parameter;
                    if (parameter == null)
                        throw new InvalidOperationException("Выберите параметр для заполнения номера.");
                    if (parameter.Id.Equals(new ElementId(BuiltInParameter.SHEET_NUMBER)))
                        throw new InvalidOperationException("Для изменения системного номера используйте столбец «Новый номер». Выберите другой параметр для заполнения.");
                }

                int startNumber = 0;
                if (SelectedMode == RenameMode.OnlyRenumber && !int.TryParse(OnlyRenumber_StartNumber, out startNumber))
                    throw new InvalidOperationException("Укажите целое стартовое значение номера листа.");
                if (SelectedMode == RenameMode.Prefix && string.IsNullOrEmpty(Prefix_Text))
                    throw new InvalidOperationException("Укажите префикс.");

                string parameterName = parameter?.Definition.Name;
                var rows = _shetsToRenumber.Select(sheet => new ListRenumberPreviewRow(sheet,
                    updateParameter ? sheet.LookupParameter(parameterName) : null) { UpdateParameter = updateParameter }).ToArray();

                if (SelectedMode == RenameMode.Unicode)
                    CalculateUnicodePreview(uidoc.Document, rows);
                else
                {
                    for (int index = 0; index < rows.Length; index++)
                    {
                        ListRenumberPreviewRow row = rows[index];
                        try
                        {
                            switch (SelectedMode)
                            {
                                case RenameMode.Prefix:
                                    row.NewNumber = ListRenumberNumberCalculator.Prefix(row.OriginalNumber, Prefix_Text);
                                    break;
                                case RenameMode.OnlyRenumber:
                                    row.NewNumber = ListRenumberNumberCalculator.Renumber(row.OriginalNumber, checked(startNumber + index));
                                    break;
                                case RenameMode.ClearedNumber:
                                    row.NewNumber = ListRenumberNumberCalculator.UserNumber(row.OriginalNumber);
                                    break;
                            }
                        }
                        catch (Exception)
                        {
                            row.CalculationError = "Не удалось рассчитать номер. Введите новый номер вручную.";
                        }
                    }
                }

                foreach (ListRenumberPreviewRow row in rows)
                    row.RefreshParameterValue();

                var preview = new ListRenumberPreviewVM(uidoc.Document, rows, SelectedMode, parameterName,
                    Unicode_SelectedUnicode?.Code, Unicode_SelectedUnicode?.Name);
                if (PreviewAction?.Invoke(preview) == true)
                    CloseAction?.Invoke();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Предпросмотр нумерации листов", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void CalculateUnicodePreview(Document document, ListRenumberPreviewRow[] rows)
        {
            if (Unicode_SelectedUnicode == null)
                throw new InvalidOperationException("Выберите символ Unicode.");
            if (!int.TryParse(Unicdoe_CountOfUnicodes, out int count) || count < 1 || count > 10000)
                throw new InvalidOperationException("Укажите количество символов Unicode от 1 до 10000.");

            var selectedIds = new HashSet<ElementId>(rows.Select(row => row.Sheet.Id));
            var occupied = new HashSet<string>(new FilteredElementCollector(document).OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>().Where(sheet => !selectedIds.Contains(sheet.Id)).Select(sheet => sheet.SheetNumber), StringComparer.Ordinal);
            var baseNumbers = new Dictionary<ListRenumberPreviewRow, string>();
            foreach (ListRenumberPreviewRow row in rows)
            {
                try
                {
                    baseNumbers[row] = Unicode_IsChangePrefixToUnicode
                        ? ListRenumberNumberCalculator.UserNumber(row.OriginalNumber) : row.OriginalNumber;
                }
                catch (Exception)
                {
                    row.CalculationError = "Не удалось рассчитать номер. Введите новый номер вручную.";
                }
            }

            // Подбираем приставку в памяти; совпадения внутри выборки пользователь увидит в таблице.
            string prefix;
            do
            {
                prefix = string.Concat(Enumerable.Repeat(Unicode_SelectedUnicode.Code, count));
                if (!baseNumbers.Values.Any(number => occupied.Contains(prefix + number))) break;
                count++;
            }
            while (count <= 10000);

            foreach (var pair in baseNumbers)
                pair.Key.NewNumber = prefix + pair.Value;
        }

        private void UpdateSelectedMode()
        {
            if (IsUnicode)
                SelectedMode = RenameMode.Unicode;
            else if (IsPrefix)
                SelectedMode = RenameMode.Prefix;
            else if (IsOnlyRenumber)
                SelectedMode = RenameMode.OnlyRenumber;
            else if (IsRefreshParam)
                SelectedMode = RenameMode.RefreshParam;
            else if (IsClearedNumber)
                SelectedMode = RenameMode.ClearedNumber;
        }



        private void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        private void Set<T>(ref T field, T value, [CallerMemberName] string prop = "")
        {
            if (!Equals(field, value))
            {
                field = value;
                OnPropertyChanged(prop);
            }
        }
    }
}
