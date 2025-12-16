using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.ExecutableCommand;
using KPLN_ExtraFilter.Forms.Commands;
using KPLN_ExtraFilter.Forms.Entities.SearchById;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_ExtraFilter.Forms.ViewModels
{
    public sealed class SearchByIdVM : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private readonly SearchByIdForm _mainWindow;

        private string _searchId;

        public SearchByIdVM(SearchByIdForm mainWindow, UIApplication uiapp, View3D special3DView)
        {
            _mainWindow = mainWindow;
            UIApp = uiapp;
            Special3DView = special3DView;
            
            Doc = uiapp.ActiveUIDocument.Document;

            SearchByIdCmd = new RelayCommand<object>(_ => SearchById(), _ => CanSearch());
            SelectByIdCmd = new RelayCommand<SearchByIdEntity>(SelectById);

            CloseWindowCmd = new RelayCommand<object>(CloseWindow);
        }

        /// <summary>
        /// Комманда: Найти элементы по ID, включая элементы из связей
        /// </summary>
        public ICommand SearchByIdCmd { get; }

        /// <summary>
        /// Комманда: Выбрать элементы по ID, включая элементы из связей
        /// </summary>
        public ICommand SelectByIdCmd { get; }

        /// <summary>
        /// Комманда: Закрыть окно
        /// </summary>
        public ICommand CloseWindowCmd { get; }

        public UIApplication UIApp { get; }

        public Document Doc { get; }
        
        /// <summary>
        /// Спец. вид для визуализации
        /// </summary>
        public View3D Special3DView { get; }

        public string SearchId
        {
            get => _searchId;
            set
            {
                _searchId = value;
                NotifyPropertyChanged();
            }
        }

        public ObservableCollection<SearchByIdEntity> SearchByIdResults { get; private set; } = new ObservableCollection<SearchByIdEntity>();

        /// <summary>
        /// Найти элементы по ID 
        /// </summary>
        private void SearchById()
        {
            // Очистка старого результата
            SearchByIdResults.Clear();

            List<SearchByIdDocEntity> searchByIdDocEntities = new List<SearchByIdDocEntity>();

            // Собираю ВСЕ элементы из линков
            IEnumerable<RevitLinkInstance> docRLIColl = new FilteredElementCollector(Doc)
                .OfClass(typeof(RevitLinkInstance))
                .WhereElementIsNotElementType()
                .Cast<RevitLinkInstance>();

            foreach (RevitLinkInstance rli in docRLIColl)
            {
                Document linkDoc = rli.GetLinkDocument();
                if (linkDoc == null)
                    continue;


                SearchByIdDocEntity searchByIdLinkDocEntity = new SearchByIdDocEntity(linkDoc, new FilteredElementCollector(linkDoc).WhereElementIsNotElementType().ToArray(), rli);
                searchByIdDocEntities.Add(searchByIdLinkDocEntity);
            }


            // Анализирую ВСЕ возможные элементы и генерирую коллекцию по элементам
            foreach (SearchByIdDocEntity sde in searchByIdDocEntities)
            {
                Element[] sdeElems = sde.GetElementsFromModelById(SearchId).Where(el => el != null).ToArray();
                if (sdeElems != null && sdeElems.Any())
                {
                    foreach (Element sdeElem in sdeElems)
                        SearchByIdResults.Add(new SearchByIdEntity(sde, sdeElem));
                }
            }
        }

        /// <summary>
        /// Метод выбора по id
        /// </summary>
        /// <param name="searchEnt"></param>
        private void SelectById(SearchByIdEntity searchEnt)
        {
            if (!UIApp.ActiveUIDocument.ActiveView.Id.Equals(Special3DView.Id) && Special3DView != null)
                UIApp.ActiveUIDocument.ActiveView = Special3DView;

            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new SelectByIdExсCmd(searchEnt));
        }

        /// <summary>
        /// Верификация ввода
        /// </summary>
        /// <returns></returns>
        private bool CanSearch() => !string.IsNullOrEmpty(SearchId) && Regex.IsMatch(SearchId, @"^\d+(\s*,\s*\d+)*$");

        public void CloseWindow(object windObj)
        {
            if (windObj is Window window)
                window.Close();
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
