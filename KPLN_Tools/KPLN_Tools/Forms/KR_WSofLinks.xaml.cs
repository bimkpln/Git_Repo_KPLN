using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace KPLN_Tools.Forms
{
    // Класс для связи (родительский элемент в TreeView)
    public class LinkWorksetsItem : INotifyPropertyChanged
    {
        //имя связи
        public string LinkName { get; set; }

        private ObservableCollection<WorksetItem> _worksets = new ObservableCollection<WorksetItem>();
        public ObservableCollection<WorksetItem> Worksets
        {
            get => _worksets;
            set
            {
                // Отписываемся от старой коллекции
                if (_worksets != null)
                {
                    foreach (var ws in _worksets)
                        ws.PropertyChanged -= OnWorksetChanged;
                }

                _worksets = value;

                // Подписываемся на новую коллекцию
                if (_worksets != null)
                {
                    foreach (var ws in _worksets)
                        ws.PropertyChanged += OnWorksetChanged;
                }

                OnPropertyChanged(nameof(Worksets));
                OnPropertyChanged(nameof(IsHighlighted));
            }
        }

        // Связь подсвечивается, если содержит хотя бы один выбранный РН
        public bool IsHighlighted => Worksets.Any(w => w.IsSelected);

        // Метод, который вызывается при изменении любого WorksetItem внутри этой связи
        private void OnWorksetChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(WorksetItem.IsSelected))
            {
                // Говорим WPF: "Эй, свойство IsHighlighted изменилось! Перерисуй круг!"
                OnPropertyChanged(nameof(IsHighlighted));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    // Класс для рабочего набора (дочерний элемент в TreeView)
    public class WorksetItem : INotifyPropertyChanged
    {
        public string Name { get; set; }
        public bool IsOpen { get; set; }
        private bool _isSelected { get; set; } // true, если есть в SelectedWorksets

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected)); // Вызываем метод который уведомляет UI об изменении свойства _isSelected
                }
            }
        }
        // Событие для уведомления UI
        public event PropertyChangedEventHandler PropertyChanged;

        // Метод для вызова события
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }

    /// <summary>
    /// Логика взаимодействия для KR_WSofLinks.xaml
    /// </summary>
    public partial class KR_WSofLinks : System.Windows.Window
    {


        //================================================================================================
        //================================================================================================
        //Переменные


        //чек-бокс управляющий сценарием работы плагина (открыть или закрыть РН)
        public bool WorksetOpenClose { get; private set; }
        private UIDocument _uidoc;

        // Все уникальные имена РН из связей (для формирования подсказок)
        public ObservableCollection<string> AllWorksetNames { get; } = new ObservableCollection<string>();

        // Отфильтрованные подсказки под введенный текст
        public ObservableCollection<string> FilteredSuggestions { get; } = new ObservableCollection<string>();

        // Выбранные пользователем рабочие наборы (в среднем окне)
        public ObservableCollection<string> SelectedWorksets { get; } = new ObservableCollection<string>();


        // Список ВСЕХ связей с рабочими наборами (в правой панели). Передается в Command_KR_WSofLinks 
        public ObservableCollection<LinkWorksetsItem> LinksWorksetsList { get; } = new ObservableCollection<LinkWorksetsItem>();

        //================================================================================================
        //================================================================================================




        private void LoadAllWorksetNamesFromLinks()
        {
            LinksWorksetsList.Clear();
            Document doc = _uidoc.Document;
            var linkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(link => (doc.GetElement(link.GetTypeId()) as RevitLinkType).IsNestedLink == false);

            var allNames = new HashSet<string>();

            foreach (RevitLinkInstance linkInstance in linkInstances)
            {
                RevitLinkType linkType = doc.GetElement(linkInstance.GetTypeId()) as RevitLinkType;
                if (linkType == null) continue;

                string linkName = linkType.Name;
                var linkItem = new LinkWorksetsItem { LinkName = linkName };

                Document linkedDoc = linkInstance.GetLinkDocument();
                // если ссылка на связь не корректная то пропускаем ее
                if (linkedDoc == null)
                    continue;

                FilteredWorksetCollector worksetCollector = new FilteredWorksetCollector(linkedDoc);
                ICollection<Workset> list_worksets = worksetCollector.ToWorksets();

                foreach (Workset ws in list_worksets)
                {
                    if (ws.Kind != WorksetKind.UserWorkset)
                        continue;

                    var worksetItem = new WorksetItem
                    {
                        Name = ws.Name,
                        IsOpen = ws.IsOpen,
                        IsSelected = false
                    };

                    worksetItem.PropertyChanged += OnWorksetPropertyChanged;

                    linkItem.Worksets.Add(worksetItem);
                    allNames.Add(ws.Name);
                }

                var sortedWorksets = linkItem.Worksets.OrderBy(w => w.Name).ToList();
                linkItem.Worksets = new ObservableCollection<WorksetItem>(sortedWorksets);

                if (!LinksWorksetsList.Contains(linkItem))
                    LinksWorksetsList.Add(linkItem);
            }

            foreach (string name in allNames.OrderBy(n => n))
                AllWorksetNames.Add(name);

            if (allNames.Count == 0)
            {
                TaskDialog.Show("Внимание",
                    "Не удалось получить рабочие наборы связей, так как связи выгружены.\n" +
                    "Рабочие наборы для недоступных связей не будут отображаться в подсказках.");
            }
        }

        private void OnWorksetNameTextChanged(object sender, TextChangedEventArgs e)
        {
            string filter = tbWorksetName.Text.Trim();

            lbSuggestions.Visibility = System.Windows.Visibility.Collapsed;
            FilteredSuggestions.Clear();

            if (string.IsNullOrEmpty(filter))
                return;

            foreach (string name in AllWorksetNames)
            {
                if (name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                    FilteredSuggestions.Add(name);
            }

            if (FilteredSuggestions.Count > 0)
                lbSuggestions.Visibility = System.Windows.Visibility.Visible;
        }

        private void OnSuggestionClick(object sender, MouseButtonEventArgs e)
        {
            if (lbSuggestions.SelectedItem is string selectedName)
            {
                if (!SelectedWorksets.Contains(selectedName))
                    SelectedWorksets.Add(selectedName);

                foreach (var link in LinksWorksetsList)
                {
                    var link_worksets = link.Worksets;
                    foreach (var workset in link_worksets)
                    {
                        if (workset.Name == selectedName)
                            workset.IsSelected = true;
                    }

                }

            }
        }

        // Двойной клик по выбранному РН → удаляем его
        private void OnSelectedWorksetClick(object sender, MouseButtonEventArgs e)
        {
            if (lbSelectedWorksets.SelectedItem is string selectedName)
            {
                SelectedWorksets.Remove(selectedName);

                foreach (var link in LinksWorksetsList)
                {
                    var link_worksets = link.Worksets;
                    foreach (var workset in link_worksets)
                    {
                        if (workset.Name == selectedName)
                            workset.IsSelected = false;
                    }

                }
            }
        }



        public KR_WSofLinks(UIDocument uidoc)
        {
            _uidoc = uidoc;
            InitializeComponent();
            LoadAllWorksetNamesFromLinks();
            DataContext = this;
            tbWorksetName.Focus();
        }

        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            SelectedWorksets.Clear();

            foreach (LinkWorksetsItem link in LinksWorksetsList)
            {
                foreach (WorksetItem workset in link.Worksets)
                {
                    workset.IsSelected = false;
                }
            }

        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {          
            //передаем режим работы (Открыть/Закрыть)
            WorksetOpenClose = chkOpenClose.IsChecked == true;
            DialogResult = LinksWorksetsList != null && LinksWorksetsList.Any();
        }

        /// <summary>
        /// Срабатывает при изменении любого свойства у WorksetItem (в т.ч. IsSelected)
        /// </summary>
        private void OnWorksetPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            // Реагируем только на изменение свойства IsSelected
            if (e.PropertyName != nameof(WorksetItem.IsSelected))
                return;

            if (sender is WorksetItem changedWorkset)
            {
                //статус изменененного чекбокса
                bool isChecked = changedWorkset.IsSelected;
                //имя РН у измененного чекбокса
                string wsName = changedWorkset.Name;


                // 🔹 ЛОГИКА ПРИ КЛИКЕ ПОЛЬЗОВАТЕЛЯ:
                if (isChecked)
                {
                    // Галку ПОСТАВИЛИ → добавляем в список выбранных (если ещё нет)
                    if (!SelectedWorksets.Contains(wsName))
                    {
                        SelectedWorksets.Add(wsName);
                    }
                }
                else
                {
                    bool IsSel = false;
                    //если сняли галку то проверяем остался ли хоть где-то данный рабочий набор выделенным
                    foreach (var link in LinksWorksetsList)
                    {
                        foreach (WorksetItem workset in link.Worksets)
                        {
                            if (workset.Name == wsName && workset.IsSelected == true)
                                IsSel = true;
                        }
                    }
                    //если не осталось то убираем из списка в центре
                    if (!IsSel)
                        SelectedWorksets.Remove(wsName);

                }
            }
        }



    }




    // =================================================================================================
    // Конвертеры для TreeView
    // =================================================================================================
    public class BoolToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return (bool)value ? Brushes.ForestGreen : Brushes.Transparent;
        }
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => throw new NotImplementedException();
    }

    public class BoolToOpenCloseConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return (bool)value ? "Открыт" : "Закрыт";
        }
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => throw new NotImplementedException();
    }

}
