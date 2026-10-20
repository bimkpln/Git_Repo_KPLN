using Autodesk.Revit.DB;
using KPLN_DefaultPanelExtension_Modify.ExecutableCommands;
using KPLN_Library_ConfigWorker;
using KPLN_Library_DBWorker.FactoryParts.SQLite;
using KPLN_Library_Forms.Services;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;

namespace KPLN_DefaultPanelExtension_Modify.Forms.Models
{
    public sealed class ListVPPositionMainVM : INotifyPropertyChanged
    {
        private readonly Window _owner;
        private const string ConfigFileName = "ViewportPositionConfig";

        // Как в Looker: ключ — полное имя файла / модели из хранилища.
        private readonly Dictionary<string, ObservableCollection<ListVPPositionCreateM>> _documentItems =
            new Dictionary<string, ObservableCollection<ListVPPositionCreateM>>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _savedDocumentData =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private Document _currentDocument;
        private string _currentDocumentKey;
        private ListVPPositionViewTypeM _selectedViewType = ListVPPositionViewTypeM.Items[0];

        public ListVPPositionMainVM(Window owner, Element[] selTrueElems)
        {
            _owner = owner;

            SelectCmd = new RelayCommand<ListVPPositionCreateM>(Select);
            CreateCmd = new RelayCommand<object>(_ => Create());
            HelpCmd = new RelayCommand<object>(_ => Help());
            DeleteCmd = new RelayCommand<ListVPPositionCreateM>(Delete);
            CloseWindowCmd = new RelayCommand<object>(CloseWindow);

            UpdateSelection(selTrueElems[0].Document, selTrueElems);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public Element[] SelectedViewElems { get; private set; } = new Element[0];

        public Element ProcessedViewElems { get; private set; }

        public ICommand SelectCmd { get; }
        public ICommand CreateCmd { get; }
        public ICommand HelpCmd { get; }
        public ICommand DeleteCmd { get; }
        public ICommand CloseWindowCmd { get; }

        public ObservableCollection<ListVPPositionCreateM> ListVPPositionCreateItems { get; private set; } =
            new ObservableCollection<ListVPPositionCreateM>();

        public ICollectionView FilteredConfigItems { get; private set; }

        public IReadOnlyList<ListVPPositionViewTypeM> ViewTypeItems => ListVPPositionViewTypeM.Items;

        public ListVPPositionViewTypeM SelectedViewType
        {
            get => _selectedViewType;
            set
            {
                if (value == null || _selectedViewType == value)
                    return;

                _selectedViewType = value;
                OnPropertyChanged();
                FilteredConfigItems?.Refresh();
            }
        }

        /// <summary>
        /// Вызывается из ExternalEvent при выборе элемента и активации вида.
        /// </summary>
        public void UpdateSelection(Document document, Element[] elements)
        {
            // Сначала исключаем применение конфигурации к устаревшей выборке.
            SelectedViewElems = new Element[0];
            ProcessedViewElems = null;

            if (document == null || !document.IsValidObject)
                return;

            string documentKey = SQLiteDocService.GetFileFullName(document);
            if (string.IsNullOrEmpty(documentKey))
                documentKey = $"{document.Title}#{document.GetHashCode()}";

            if (!string.Equals(documentKey, _currentDocumentKey, StringComparison.OrdinalIgnoreCase))
            {
                SaveConfig();

                if (!_documentItems.TryGetValue(documentKey, out ObservableCollection<ListVPPositionCreateM> items))
                {
                    var loadedItems = ConfigService.ReadConfigFile<ListVPPositionCreateM[]>(
                        ModuleData.RevitVersion, document, ConfigType.Shared, ConfigFileName) as ListVPPositionCreateM[]
                        ?? new ListVPPositionCreateM[0];
                    items = new ObservableCollection<ListVPPositionCreateM>(
                        loadedItems.Where(item => item != null).OrderBy(item => item.ConfigName));

                    _documentItems.Add(documentKey, items);
                    _savedDocumentData.Add(documentKey, GetItemsData(items));
                }

                _currentDocumentKey = documentKey;
                ListVPPositionCreateItems = items;
                FilteredConfigItems = new ListCollectionView(items);
                FilteredConfigItems.Filter = MatchesViewType;
                FilteredConfigItems.SortDescriptions.Add(
                    new SortDescription(nameof(ListVPPositionCreateM.ConfigName), ListSortDirection.Ascending));

                OnPropertyChanged(nameof(ListVPPositionCreateItems));
                OnPropertyChanged(nameof(FilteredConfigItems));
            }

            _currentDocument = document;
            SelectedViewElems = (elements ?? new Element[0])
                .Where(element => element != null && element.IsValidObject && element.Document.Equals(document))
                .ToArray();

            if (SelectedViewElems.Length == 1)
            {
                ProcessedViewElems = SelectedViewElems[0];
                string code = ListVPPositionViewTypeM.GetCode(ProcessedViewElems);
                bool hasMatchingConfigs = ListVPPositionCreateItems.Any(
                    item => ListVPPositionViewTypeM.NormalizeCode(item.ViewTypeCode) == code);

                // Если конфигураций этого типа ещё нет, показываем в том числе легаси без типа.
                SelectedViewType = hasMatchingConfigs
                    ? ViewTypeItems.First(item => item.Code == code)
                    : ViewTypeItems[0];
            }
            else
                SelectedViewType = ViewTypeItems[0];
        }

        public void SaveConfig()
        {
            if (_currentDocument == null || !_currentDocument.IsValidObject || _currentDocumentKey == null)
                return;

            string data = GetItemsData(ListVPPositionCreateItems);
            if (_savedDocumentData[_currentDocumentKey] == data)
                return;

            // Сохраняем всю коллекцию документа; фильтр влияет только на отображение.
            ConfigService.SaveConfig<ListVPPositionCreateM>(
                ModuleData.RevitVersion, _currentDocument, ConfigType.Shared, ListVPPositionCreateItems, ConfigFileName);
            _savedDocumentData[_currentDocumentKey] = data;
        }

        private bool MatchesViewType(object value)
        {
            if (!(value is ListVPPositionCreateM item))
                return false;

            return string.IsNullOrEmpty(SelectedViewType.Code)
                || ListVPPositionViewTypeM.NormalizeCode(item.ViewTypeCode) == SelectedViewType.Code;
        }

        private bool CheckANDSetProcessedViewport()
        {
            if (SelectedViewElems.Length != 1 || _currentDocument == null || !_currentDocument.IsValidObject)
            {
                MessageBox.Show(_owner, "Выбери один видовой экран или одну спецификацию на листе.",
                    "Положение вида", MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }

            Element selected = SelectedViewElems[0];
            if (!selected.IsValidObject || !selected.Document.Equals(_currentDocument))
                return false;

            ProcessedViewElems = selected;
            return true;
        }

        private void Select(ListVPPositionCreateM item)
        {
            if (item != null && ListVPPositionCreateItems.Contains(item) && CheckANDSetProcessedViewport())
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ExcCmdListVPPositionSet(ProcessedViewElems, item));
        }

        private void Create()
        {
            if (!CheckANDSetProcessedViewport())
                return;

            ListVPPositionCreateFrom createFrom = new ListVPPositionCreateFrom(_owner, ProcessedViewElems);
            WindowHandleSearch.MainWindowHandle.SetAsOwner(createFrom);

            if (createFrom.ShowDialog() == true)
            {
                ListVPPositionCreateItems.Add(createFrom.CurentListVPPositionCreateVM.ListVPPositionItem);
                SaveConfig();
            }
        }

        private void Help() =>
            Process.Start(new ProcessStartInfo(ExcCmdListVPPositionStart.HelpUrl) { UseShellExecute = true });

        private void Delete(ListVPPositionCreateM item)
        {
            if (item == null || !ListVPPositionCreateItems.Contains(item))
                return;

            var result = MessageBox.Show(_owner, $"Будет удалена конфигурация \"{item.ConfigName}\".",
                "Внимание", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                ListVPPositionCreateItems.Remove(item);
                SaveConfig();
            }
        }

        private void CloseWindow(object windObj)
        {
            if (windObj is Window window)
                window.Close();
        }

        private static string GetItemsData(IEnumerable<ListVPPositionCreateM> items) =>
            JsonConvert.SerializeObject(items.Select(item => item.ToJson()));

        private void OnPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
