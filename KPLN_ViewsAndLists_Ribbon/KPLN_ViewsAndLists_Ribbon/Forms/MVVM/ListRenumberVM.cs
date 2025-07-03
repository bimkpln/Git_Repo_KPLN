using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_PluginActivityWorker;
using KPLN_ViewsAndLists_Ribbon.Common.Lists;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;

namespace KPLN_ViewsAndLists_Ribbon.Forms.MVVM
{
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

        private Parameter _prefix_SelectedTitleBlockParameter;
        private Parameter _onlyRenumber_SelectedTitleBlockParameter;
        private Parameter _refreshParam_SelectedTitleBlockParameter;

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
            TitleBlockParameters = new ObservableCollection<Parameter>(tBlockParams);

            OkCommand = new RelayCommand(OnOk);
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

        public ObservableCollection<Parameter> TitleBlockParameters { get; set; }

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

        public Parameter Prefix_SelectedTitleBlockParameter
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

        public Parameter OnlyRenumber_SelectedTitleBlockParameter
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

        public Parameter RefreshParam_SelectedTitleBlockParameter
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
        public RelayCommand OkCommand { get; set; }

        public RelayCommand CancelCommand { get; set; }

        public Action CloseAction { get; set; }

        private void OnOk()
        {
            Document doc = _uiapp.ActiveUIDocument.Document;
            UIDocument uidoc = _uiapp.ActiveUIDocument;
            if (uidoc == null) return;

            CloseAction?.Invoke();
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName($"{PluginName}", ModuleData.ModuleName).ConfigureAwait(false);

            using (Transaction trans = new Transaction(doc, $"KPLN: {PluginName}"))
            {
                trans.Start();
                try
                {
                    switch (SelectedMode)
                    {
                        // Меняю номер листа с использованием символов Юникода
                        case RenameMode.Unicode:
                            UniEntity cmbSelUni = Unicode_SelectedUnicode;
                            if (cmbSelUni == null)
                            {
                                MessageBox.Show("Вы забили выбрать Юникод", "Ошибка");
                                return;
                            }

                            int uniNumber = Int32.Parse(Unicdoe_CountOfUnicodes) - 1;
                            bool isRun = true;
                            while (isRun)
                            {
                                uniNumber++;
                                isRun = UseUniCodes(_shetsToRenumber, cmbSelUni, uniNumber, Unicode_IsChangePrefixToUnicode);
                            }
                            break;

                        // Меняю номер листа с использованием префиксов
                        case RenameMode.Prefix:
                            if (string.IsNullOrEmpty(Prefix_Text))
                            {
                                MessageBox.Show("Вы забили указать префикс", "Ошибка");
                                return;
                            }
                            string userPrefix = $"{Prefix_Text}/";
                            UsePrefix(_shetsToRenumber, userPrefix);

                            if (Prefix_IsParamUpdate)
                                ParamRefresh(_shetsToRenumber, Prefix_SelectedTitleBlockParameter);

                            break;

                        //Меняю номер листа без редактирования префиксов и Юникодов
                        case RenameMode.OnlyRenumber:
                            if (int.TryParse(OnlyRenumber_StartNumber, out int or_startNumber))
                            {
                                ClearRenumber(_shetsToRenumber, or_startNumber);

                                if (OnlyRenumber_IsParamUpdate)
                                    ParamRefresh(_shetsToRenumber, OnlyRenumber_SelectedTitleBlockParameter);
                            }

                            break;

                        // Заполняю пользовательский параметр для нумерации в штампе
                        case RenameMode.RefreshParam:
                            ParamRefresh(_shetsToRenumber, RefreshParam_SelectedTitleBlockParameter);
                            break;

                        // Очистка от приставок и юникодов
                        case RenameMode.ClearedNumber:
                            ClearedNumber(_shetsToRenumber);
                            break;
                    }

                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("Sheet number is already in use"))
                    {
                        TaskDialog.Show("Предупреждение", "Такой номер листа уже есть. Работа экстренно завершена!", TaskDialogCommonButtons.Ok);
                        trans.RollBack();
                        return;
                    }
                    else
                    {
                        TaskDialog.Show("Ошибка", $"Отправь в BIM-отдел\n\n{ex.StackTrace}\n{ex.Message}\n\nРабота экстренно завершена!", TaskDialogCommonButtons.Ok);
                        trans.RollBack();
                        return;
                    }
                }
                trans.Commit();

                // Обновляю нумерацию в браузере проекта
                DockablePaneId dpId = DockablePanes.BuiltInDockablePanes.ProjectBrowser;
                DockablePane dP = new DockablePane(dpId);
                dP.Show();
            }
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



        /// <summary>
        /// Метод для замены номер листа с использованием символов Юникода
        /// </summary>
        /// <example> 
        /// "7.1" преобразуется в "символЮникода7.1"
        /// </example>
        private static bool UseUniCodes(IEnumerable<ViewSheet> sortedSheets, UniEntity cmbSelUni, int counter, bool isConvertToUni)
        {
            try
            {
                foreach (ViewSheet curVSheet in sortedSheets)
                {
                    string trueNumbStr;
                    if (isConvertToUni)
                    {
                        string strVSheetNumb = curVSheet.SheetNumber;
                        trueNumbStr = UserNumber(strVSheetNumb);
                    }
                    else
                        trueNumbStr = curVSheet.SheetNumber;

                    curVSheet.SheetNumber = String.Concat(Enumerable.Repeat(cmbSelUni.Code, counter)) + trueNumbStr;

                    Parameter uniCodeParam = curVSheet.get_Parameter(new Guid("09b934d4-81b3-4aff-b37e-d20dfbc1ac8e"));
                    if (uniCodeParam != null)
                    {
                        if (counter == 1)
                            curVSheet.get_Parameter(new Guid("09b934d4-81b3-4aff-b37e-d20dfbc1ac8e")).Set(cmbSelUni.Name);
                        else
                            curVSheet.get_Parameter(new Guid("09b934d4-81b3-4aff-b37e-d20dfbc1ac8e")).Set($"{counter}X{cmbSelUni.Name}");
                    }
                }
                return false;
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {
                return true;
            }
        }

        /// <summary>
        /// Метод преобразования номера с символами в число.
        /// <example> 
        /// Например: АР1/007.1 преобразуется в 7.1
        /// </example>
        /// </summary>
        private static string UserNumber(string number)
        {
            char[] charArray = number.ToCharArray();
            foreach (char c in charArray)
            {
                if (Char.IsLetter(c))
                {
                    number = number.Trim(c);
                }
                else if (Enum.IsDefined(typeof(UniDecCodes), (int)c))
                {
                    number = number.Trim(c);
                }
                else if (c.Equals('.'))
                {
                    continue;
                }
                else if (Char.IsPunctuation(c))
                {
                    number = number.Split(c)[1];
                }
            }
            if (number.All(c => c.Equals('0'))) return "0";
            else number = number.Length > 1 ? number.TrimStart('0') : number;

            return number;
        }

        /// <summary>
        /// Метод для замены номер листа с использованием символов префикса
        /// </summary>
        /// <example> 
        /// "7.1" преобразуется в "АР/7.1", или в "007.1"
        /// </example>
        private static void UsePrefix(IEnumerable<ViewSheet> sortedSheets, string userPrefix)
        {
            foreach (ViewSheet curVSheet in sortedSheets)
            {
                string strVSheetNumb = curVSheet.SheetNumber;
                string onlyNumb = UserNumber(strVSheetNumb);

                string ZeroNumber;
                if (OnlyNumber(onlyNumb) < 10)
                {
                    ZeroNumber = $"00{onlyNumb}";
                }
                else if (OnlyNumber(onlyNumb) < 100)
                {
                    ZeroNumber = $"0{onlyNumb}";
                }
                else
                {
                    ZeroNumber = $"{onlyNumb}";
                }
                curVSheet.SheetNumber = userPrefix + ZeroNumber;
            }
        }

        /// <summary>
        /// Метод преобразования номера с символами в число без подразделов
        /// </summary>
        /// <example> 
        /// АР/007.1 преобразуется в 7
        /// </example>
        private static int OnlyNumber(string number)
        {
            char[] charArray = number.ToCharArray();
            foreach (char c in charArray)
            {
                if (c.Equals('.'))
                {
                    number = number.Split('.')[0];
                    break;
                }
            }
            return Int32.Parse(new String(number.Where(Char.IsDigit).ToArray()));
        }

        /// <summary>
        /// Метод для заполнения номером пользовательского параметра
        /// </summary>
        /// <example> 
        /// "АР/7.1" заполниться в выбранный параметр как "7.1"
        /// </example>
        private static void ParamRefresh(IEnumerable<ViewSheet> sortedSheets, Parameter cmbSelPar)
        {
            foreach (ViewSheet curVSheet in sortedSheets)
            {
                string strVSheetNumb = curVSheet.SheetNumber;
                string cmbSelParName = cmbSelPar.Definition.Name;
                string onlyNumb = UserNumber(strVSheetNumb);
                if (Int32.TryParse(strVSheetNumb, out int refreshNumber))
                    curVSheet.LookupParameter(cmbSelParName).Set(refreshNumber.ToString());
                else
                    curVSheet.LookupParameter(cmbSelParName).Set(onlyNumb);
            }
        }

        /// <summary>
        /// Метод для замены номер листа с использованием стартового номера и с сохранением приставок и т.п.
        /// </summary>
        /// <example> 
        /// "АР/007.1" преобразуется в "АР/008.1"
        /// </example>
        private static void ClearRenumber(ViewSheet[] sortedSheets, int startNumber)
        {
            // Получаю стартовую разницу между номерами
            int startVSheetNumber = OnlyNumber(sortedSheets[0].SheetNumber);
            int deltaNumber = startNumber - startVSheetNumber;

            ViewSheet[] reversedSheets = sortedSheets.Reverse().ToArray();
            // Задаю нумерацию с учетом стартовой разницы
            foreach (ViewSheet curVSheet in reversedSheets)
            {
                string uniCode = string.Empty;
                string textPrefix = string.Empty;
                string zeroNumberPrefix = string.Empty;
                int number = OnlyNumber(curVSheet.SheetNumber) + deltaNumber;
                string subNumberOfNumber = string.Empty;

                string sheetNumber = curVSheet.SheetNumber;
                var matchNumberByPrefix = Regex.Match(sheetNumber, @"([А-ЯA-Z]+)/");
                var matchNumberByPrefixWithNumb = Regex.Match(sheetNumber, @"([А-ЯA-Z]+)\d/");
                var matchByZeroPrefix = Regex.Match(sheetNumber, @"(/[0]?)");
                var matchBySubNumber = Regex.Match(sheetNumber, @"(\d?)\.(\d?)");

                if (matchNumberByPrefix.Success)
                    textPrefix = matchNumberByPrefix.Value;

                if (matchNumberByPrefixWithNumb.Success)
                    textPrefix = matchNumberByPrefixWithNumb.Value;

                if (matchByZeroPrefix.Success)
                {
                    if (number < 10)
                        zeroNumberPrefix = "00";
                    else if (number < 100)
                        zeroNumberPrefix = "0";
                }

                if (matchBySubNumber.Success)
                    subNumberOfNumber = $".{matchBySubNumber.Value.Split('.')[1]}";

                curVSheet.SheetNumber = $"{textPrefix}{zeroNumberPrefix}{number}{subNumberOfNumber}";
            }
        }

        /// <summary>
        /// Метод очистки номера от приставок и юникодов
        /// </summary>
        private static void ClearedNumber(ViewSheet[] sortedSheets)
        {
            foreach (ViewSheet vs in sortedSheets)
            {
                string strVSheetNumb = vs.SheetNumber;
                vs.SheetNumber = UserNumber(strVSheetNumb);
                Parameter uniCodeParam = vs.get_Parameter(new Guid("09b934d4-81b3-4aff-b37e-d20dfbc1ac8e"));
                if (uniCodeParam != null)
                    vs.get_Parameter(new Guid("09b934d4-81b3-4aff-b37e-d20dfbc1ac8e")).Set(string.Empty);
            }
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