using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using KPLN_ExtraFilter.Common;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Controls;

namespace KPLN_ExtraFilter.Forms.Entities
{
    /// <summary>
    /// Фильтры для выбора элементов в модели, или на виде (использовать с RadioButton). 
    /// Bool - не подходит, т.к. WPF Framework под капотом сетит значения, чтобы сохранить работу RadioButton одной группы
    /// </summary>
    public enum ViewFilterMode
    {
        CurrentView,
        Model
    }

    /// <summary>
    /// Фильтры для выбора элементов - новая выборка, или добавить к текущей (использовать с RadioButton). 
    /// Bool - не подходит, т.к. WPF Framework под капотом сетит значения, чтобы сохранить работу RadioButton одной группы
    /// </summary>
    public enum SelectFilterMode
    {
        CreateNew,
        AddCurrent
    }

    /// <summary>
    /// Модель для SelectionByModel
    /// </summary>
    public sealed class SelectionByModelM : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private View _docActiveView;
        private IEnumerable<Element> _userSelElems;
        private Element[] _where_UserSelElems;
        private bool _only3D;
        private bool _where_Workset;
        private WSEntity _where_SelectedWorkset;
        private bool _whereCategoryFilter;
        private bool _what_ParameterData;
        private ViewFilterMode _where_ViewDocFilterMode = ViewFilterMode.CurrentView;
        private SelectFilterMode _how_SelectFilterMode = SelectFilterMode.CreateNew;
        private bool _belongGroup;

        private TreeElementEntity[] _treeElemEntities;

        private bool _isWorkshared = false;
        private bool _canRun = false;
        private string _userHelp;


        [JsonConstructor]
        public SelectionByModelM() { }

        public SelectionByModelM(Document doc)
        {
            Doc = doc;
            IsWorkshared = doc.IsWorkshared || doc.IsDetached;
            DocActiveView = doc.ActiveView;
        }

        public Document Doc { get; set; }

        public View DocActiveView
        {
            get => _docActiveView;
            set
            {
                if (value != _docActiveView && Where_ViewDocFilterMode == ViewFilterMode.CurrentView)
                {
                    _docActiveView = value;
                    
                    SetUserSelElems();
                    UpdateCanRunANDUserHelp();
                }
            }
        }

        /// <summary>
        /// Пользовательский выбор из модели
        /// </summary>
        public IEnumerable<Element> UserSelElems { get => _userSelElems; set => _userSelElems = value; }

        /// <summary>
        /// Коллекция элементов согласно пользовательскому выбору
        /// </summary>
        public Element[] Where_UserSelElems
        {
            get => _where_UserSelElems;
            private set
            {
                if (_where_UserSelElems == null || !AreElementsEqual(_where_UserSelElems, value))
                {
                    _where_UserSelElems = value;
                    NotifyPropertyChanged(nameof(Where_UserSelCount));
                }
            }
        }

        /// <summary>
        /// Коллекция элементов БЕЗ фильтрации по категории
        /// </summary>
        public Element[] Cahce_UserSelElemsWithoutCatFilter { get; set; }

        /// <summary>
        /// На виде/в модели
        /// </summary>
        public ViewFilterMode Where_ViewDocFilterMode
        {
            get => _where_ViewDocFilterMode;
            set
            {
                _where_ViewDocFilterMode = value;
                NotifyPropertyChanged();

                if (Where_Category)
                    ReloadCategories();
                if (What_ParameterData)
                    ReloadParams();

                SetUserSelElems();
                UpdateCanRunANDUserHelp();
            }
        }

        /// <summary>
        /// Только 3D
        /// </summary>
        public bool Where_Only3D
        {
            get => _only3D;
            set
            {
                _only3D = value;
                NotifyPropertyChanged();

                SetUserSelElems();
                UpdateCanRunANDUserHelp();
            }
        }

        /// <summary>
        /// Одного рабочего набора
        /// </summary>
        public bool Where_Workset
        {
            get => _where_Workset;
            set
            {
                _where_Workset = value;
                NotifyPropertyChanged();
                NotifyPropertyChanged(nameof(Where_DocWorksets));
                WorksetFilter.LoadCollection(Where_DocWorksets);
                NotifyPropertyChanged(nameof(FilteredWSView));

                SetUserSelElems();
                UpdateCanRunANDUserHelp();
            }
        }

        /// <summary>
        /// Экземпляр фильтра по РН
        /// </summary>
        public CollectionFilter<WSEntity> WorksetFilter { get; } = new CollectionFilter<WSEntity>();

        /// <summary>
        /// Пользовательский ввод для поиска РН
        /// </summary>
        public string SearchWSText
        {
            get => WorksetFilter.SearchText;
            set => WorksetFilter.SearchText = value;
        }

        /// <summary>
        /// Коллекция РН для окна
        /// </summary>
        public ICollectionView FilteredWSView => WorksetFilter.View;

        /// <summary>
        /// Коллекция рабочих наборов модели
        /// </summary>
        public WSEntity[] Where_DocWorksets
        {
            get
            {
                if (!Where_Workset && (Doc == null || !IsWorkshared))
                    return null;

                return new FilteredWorksetCollector(Doc)
                    .OfKind(WorksetKind.UserWorkset)
                    .OrderBy(ws => ws.Name)
                    .Select(ws => new WSEntity(ws))
                    .ToArray();
            }
        }

        /// <summary>
        /// Выбранный пользователем РН
        /// </summary>
        public WSEntity Where_SelectedWorkset
        {
            get => _where_SelectedWorkset;
            set
            {
                _where_SelectedWorkset = value;
                NotifyPropertyChanged();

                SetUserSelElems();
                UpdateCanRunANDUserHelp();
            }
        }

        /// <summary>
        /// Выбранной категории/-ий
        /// </summary>
        public bool Where_Category
        {
            get => _whereCategoryFilter;
            set
            {
                _whereCategoryFilter = value;
                NotifyPropertyChanged();

                if (_whereCategoryFilter)
                    Where_SelectedCategories.Add(new SelectionByModelM_CategoryM(this));
                else
                    Where_SelectedCategories.Clear();

                SetUserSelElems();
                UpdateCanRunANDUserHelp();
            }
        }

        /// <summary>
        /// Коллекция выбранных категорий для использования в фильтрах
        /// </summary>
        public ObservableCollection<SelectionByModelM_CategoryM> Where_SelectedCategories { get; } = new ObservableCollection<SelectionByModelM_CategoryM>();

        /// <summary>
        /// Исключить элементы групп
        /// </summary>
        public bool Belong_Group
        {
            get => _belongGroup;
            set
            {
                _belongGroup = value;
                NotifyPropertyChanged();

                SetUserSelElems();
                UpdateCanRunANDUserHelp();
            }
        }

        /// <summary>
        /// Группировать по значению параметра
        /// </summary>
        public bool What_ParameterData
        {
            get => _what_ParameterData;
            set
            {
                _what_ParameterData = value;
                NotifyPropertyChanged();

                if (_what_ParameterData)
                    What_SelectedParameters.Add(new SelectionByModelM_ParamM(this));
                else
                    What_SelectedParameters.Clear();

                SetUserSelElems();
                UpdateCanRunANDUserHelp();
            }
        }

        /// <summary>
        /// Коллекция выбранных параметров для использования в группировании
        /// </summary>
        public ObservableCollection<SelectionByModelM_ParamM> What_SelectedParameters { get; } = new ObservableCollection<SelectionByModelM_ParamM>();

        /// <summary>
        /// Создать новую выборку/Добавить выборку к уже выделенным элементам в модели
        /// </summary>
        public SelectFilterMode How_SelectFilterMode
        {
            get => _how_SelectFilterMode;
            set
            {
                _how_SelectFilterMode = value;
                NotifyPropertyChanged();

                UpdateCanRunANDUserHelp();                
            }
        }

        /// <summary>
        /// Коллекция элементов в дереве
        /// </summary>
        public TreeElementEntity[] TreeElemEntities
        {
            get => _treeElemEntities;
            set
            {
                _treeElemEntities = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Маркер модели из хранилища
        /// </summary>
        public bool IsWorkshared
        {
            get => _isWorkshared;
            set
            {
                _isWorkshared = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Маркер возможности запуска
        /// </summary>
        public bool CanRun
        {
            get => _canRun;
            set
            {
                _canRun = value;
                NotifyPropertyChanged();

                if (_canRun)
                    CreateTree();
            }
        }

        /// <summary>
        /// Пользовательская подсказка
        /// </summary>
        public string UserHelp
        {
            get => _userHelp;
            set
            {
                _userHelp = string.IsNullOrWhiteSpace(value) ? string.Empty : $"ВАЖНО: {value}";
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Подсчёт элементов в окне
        /// </summary>
        public string Where_UserSelCount => $"Под условия фильтра попало {Where_UserSelElems.Length} эл.";

        /// <summary>
        /// Устанавливаю значения по умолчанию (для сброса сущности)
        /// </summary>
        /// <returns></returns>
        public SelectionByModelM DropToDefault()
        {
            Where_ViewDocFilterMode = ViewFilterMode.CurrentView;
            Where_Workset = false;
            Where_Category = false;
            Belong_Group = false;
            What_ParameterData = false;
            How_SelectFilterMode = SelectFilterMode.CreateNew;

            return this;
        }

        /// <summary>
        /// Обновить статус возможности запуска и текстовых подсказок
        /// </summary>
        public void UpdateCanRunANDUserHelp()
        {
            // Обновляю статус ВОЗМОЖНОСТИ запуска
            bool checkElems = Where_UserSelElems != null && Where_UserSelElems.Length > 0;

            bool checkWs = true;
            if (Where_Workset && Where_SelectedWorkset == null)
                checkWs = false;

            bool checkCategory = true;
            if (Where_Category && Where_SelectedCategories.Count == 0 || Where_SelectedCategories.Any(catM => catM.CatM_SelectedCategory == null))
                checkCategory = false;

            bool checkParam = true;
            if (What_ParameterData && What_SelectedParameters.Count == 0 || What_SelectedParameters.Any(paramM => paramM.ParamM_SelectedParameter == null))
                checkParam = false;

            CanRun = checkElems && checkWs && checkCategory && checkParam;



            // Обновляю комментарий-подсказку пользователю
            if (Where_Workset && Where_SelectedWorkset == null)
            {
                UserHelp = "Выбери рабочий набор для анализа, или сними галку с фильтрации по рабочим наборам";
                return;
            }

            if (Where_Category && Where_SelectedCategories.Any(catM => catM.CatM_SelectedCategory == null))
            {
                UserHelp = "Одна или несколько категорий не заполнены. Или заполни, или сними галку с фильтрации по категориям";
                return;
            }

            if (What_ParameterData && What_SelectedParameters.Any(paramM => paramM.ParamM_SelectedParameter == null))
            {
                UserHelp = "Один или несколько параметров не заполнены. Или заполни, или сними галку с группирования по параметрам";
                return;
            }

            if (Where_UserSelElems == null || Where_UserSelElems.Length == 0)
            {
                UserHelp = "Ничего не попало в выборку. Проверь блок \"Откуда анализ?\"";
                return;
            }

            UserHelp = string.Empty;
        }

        /// <summary>
        /// Установить/обновить пользовательский выбор элементов
        /// </summary>
        public void SetUserSelElems()
        {
            // FIC по виду/документу
            FilteredElementCollector fic = null;
            switch (Where_ViewDocFilterMode)
            {
                case ViewFilterMode.CurrentView:
                    fic = new FilteredElementCollector(Doc, Doc.ActiveView.Id);
                    break;
                case ViewFilterMode.Model:
                    fic = new FilteredElementCollector(Doc);
                    break;
            }

            if (fic == null)
                return;

            // Выкл типы
            fic = fic.WhereElementIsNotElementType();

            // Поиск по рабочему набору
            if (Where_Workset && Where_SelectedWorkset != null)
            {
                ElementFilter sameTypeFilter = SelectionSearchFilter.SearchByWSEntity(Where_SelectedWorkset);
                fic = fic.WherePasses(sameTypeFilter);
            }

            List<Element> elemsNoCat = fic.Where(el => el.Category != null).ToList();


            // Поиск по значению параметра
            IEnumerable<Element> elemsWithCat = elemsNoCat;
            if (Where_Category && Where_SelectedCategories.All(c => c.CatM_SelectedCategory != null))
            {
                List<ElementFilter> catFilters = Where_SelectedCategories
                    .Select(c => SelectionSearchFilter.SearchByCategoryEntity(c.CatM_SelectedCategory))
                    .Cast<ElementFilter>()
                    .ToList();

                if (catFilters.Count > 0)
                {
                    var orFilter = new LogicalOrFilter(catFilters);
                    elemsWithCat = fic.WherePasses(orFilter).Where(e => e.Category != null);
                }
            }


            // Фильтр по 3Д
            if (Where_Only3D)
            {
                elemsNoCat = elemsNoCat.Where(el => el.Category.CategoryType == CategoryType.Model).ToList();
                elemsWithCat = elemsWithCat.Where(el => el.Category.CategoryType == CategoryType.Model);
            }


            // Исключаю группы
            if (Belong_Group)
            {
                elemsNoCat = elemsNoCat.Where(el => el.GroupId.Equals(ElementId.InvalidElementId)).ToList();
                elemsWithCat = elemsWithCat.Where(el => el.GroupId.Equals(ElementId.InvalidElementId));
            }


            Where_UserSelElems = elemsWithCat.ToArray();
            Cahce_UserSelElemsWithoutCatFilter = elemsNoCat.ToArray();
            
            CreateTree();
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        /// <summary>
        /// Сгенерировать дерево в окне
        /// </summary>
        private void CreateTree()
        {
            //Группировка по параметру
            if (What_ParameterData && What_SelectedParameters.All(paramM => paramM.ParamM_SelectedParameter != null))
                TreeElemEntities = TreeElementEntity.CreateTreeElEnt_ByParamANDCatANDFamANDType(Doc, Where_UserSelElems, What_SelectedParameters);
            // Группировка по категории
            else
                TreeElemEntities = TreeElementEntity.CreateTreeElEnt_ByCatANDFamANDType(Where_UserSelElems);
        }

        /// <summary>
        /// Сравнить коллекции элементов по id
        /// </summary>
        private bool AreElementsEqual(Element[] a, Element[] b)
        {
            if (a == null || b == null)
                return false;

            if (a.Length != b.Length)
                return false;

            // Сравнение по ElementId
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            return a.Select(e => e.Id.IntegerValue)
                    .OrderBy(id => id)
                    .SequenceEqual(b.Select(e => e.Id.IntegerValue).OrderBy(id => id));
#else
            return a.Select(e => e.Id.Value)
                    .OrderBy(id => id)
                    .SequenceEqual(b.Select(e => e.Id.Value).OrderBy(id => id));
#endif
        }

        /// <summary>
        /// Перезагружает список категорий
        /// </summary>
        private void ReloadCategories()
        {
            if (Where_Category)
            {
                foreach (var catM in Where_SelectedCategories)
                {
                    var tempOldSelCatM = catM.CatM_SelectedCategory;
                    catM.CatM_UserSelElems = Cahce_UserSelElemsWithoutCatFilter;

                    if (tempOldSelCatM != null)
                        catM.RestoreSelectedCategoryByName(tempOldSelCatM.RevitCatName);
                }
            }
        }

        /// <summary>
        /// Перезагружает список параметров
        /// </summary>
        private void ReloadParams()
        {
            if (UserSelElems == null)
                return;
            
            if (What_ParameterData)
            {
                foreach (var paramM in What_SelectedParameters)
                {
                    var tempOldSelParamM = paramM.ParamM_SelectedParameter;
                    paramM.ParamM_UserSelElems = UserSelElems.ToArray();

                    if (tempOldSelParamM != null)
                        paramM.RestoreSelectedParamById(tempOldSelParamM.RevitParamIntId);
                }
            }
        }
    }
}