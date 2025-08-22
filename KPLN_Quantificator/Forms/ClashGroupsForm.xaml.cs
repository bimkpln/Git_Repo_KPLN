using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Interop;
using KPLN_Quantificator.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static KPLN_Quantificator.Common.Collections;
using Application = Autodesk.Navisworks.Api.Application;

namespace KPLN_Quantificator.Forms
{
    /// <summary>
    /// Логика взаимодействия для ClashGroupsForm.xaml
    /// </summary>
    public partial class ClashGroupsForm : Window
    {
        public ClashGroupsForm()
        {
            InitializeComponent();
            RegisterChanges();
            
            DataContext = this;

            this.PreviewKeyDown += ClashGroupsForm_PreviewKeyDown;

            string currentDisplayName = ClashCurrentIssue.CurrentTest?.DisplayName;

            if (!string.IsNullOrEmpty(currentDisplayName))
            {
                SearchText.Text = currentDisplayName;
                this.Dispatcher.InvokeAsync(() =>
                {
                    var list = ClashTestListBox.ItemsSource as IEnumerable<CustomClashTest>;
                    if (list != null)
                    {
                        var match = list.FirstOrDefault(x => x.DisplayName == currentDisplayName);
                        if (match != null)
                        {
                            ClashTestListBox.SelectedItem = match;
                            ClashTestListBox.ScrollIntoView(match);
                        }
                    }

                    // Ключевые слова для SelectionA
                    string[] selectionAKeys = {
                        "Стены", "Витраж", "Фасад", "Окна", "Двери", "Ворота", "Перемычки", "Колонны", "Каркас"
                    };
                    // Ключевые слова для GridIntersection
                    string[] selectionGrindKeys = {
                        "Потолки", "Кровля", "Фасад", "Эвакуац", "Объёмы", "Перекрытия", "Фундамент"
                    };
                    // Ключевые слова для Main
                    string[] selectionMainKeys = {
                        "Воздуховоды", "Трубы", "Лотки", "Оборудование", "Электрооборудование", "Шахты"
                    };

                    string name = currentDisplayName?.Trim() ?? "";

                    // Получаем количество конфликтов
                    int clashCount = ClashCurrentIssue.CurrentTest?.Children
                        .OfType<ClashResult>()
                        .Count() ?? 0;

                    // Определяем режим группировки
                    GroupingMode mode;

                    if (selectionMainKeys.Any(key => name.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        if (selectionAKeys.Any(key => name.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            mode = GroupingMode.SelectionA;
                        }
                        else if (selectionGrindKeys.Any(key => name.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            mode = GroupingMode.GridIntersection;
                        }
                        else
                        {
                            if (clashCount < 50)
                            {
                                mode = GroupingMode.Host;
                            }
                            else
                            {
                                mode = GroupingMode.GridIntersection;
                            }
                        }
                    }
                    else
                    {
                        if (clashCount < 50)
                        {
                            mode = GroupingMode.Host;
                        }
                        else
                        {
                            mode = GroupingMode.GridIntersection;
                        }
                    }

                    // Устанавливаем выбранный режим
                    comboBoxGroupBy.SelectedItem = mode;
                }, System.Windows.Threading.DispatcherPriority.Loaded);
            }
        }

        public ObservableCollection<GroupingMode> GroupByList { get; set; } = new ObservableCollection<GroupingMode>();
        
        public ObservableCollection<GroupingMode> GroupThenList { get; set; } = new ObservableCollection<GroupingMode>();
        
        public ObservableCollection<CustomClashTest> ClashTests { get; set; } = new ObservableCollection<CustomClashTest>();
        //public ClashTest SelectedClashTest { get; set; }
        
        public void GetClashTests()
        {
            DocumentClashTests dct = Application.MainDocument.GetClash().TestsData;
            ClashTests.Clear();
            foreach (SavedItem savedItem in dct.Tests)
            {
                if (savedItem.GetType() == typeof(ClashTest))
                {
                    ClashTests.Add(new CustomClashTest(savedItem as ClashTest));
                }
            }
        }
        









        private void Group_Button_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            string clashStatus = radioNewActive.IsChecked == true ? "NewActiveClash" : "AllClash";

            if (ClashTestListBox.SelectedItems.Count != 0)
            {
                foreach (object selectedItem in ClashTestListBox.SelectedItems)
                {
                    CustomClashTest selectedClashTest = (CustomClashTest)selectedItem;

                    ClashTest clashTestOriginal = selectedClashTest.ClashTest;

                    try
                    {                      
                        if (clashTestOriginal.Children.Count != 0)
                        {

                            if (comboBoxGroupBy.SelectedItem == null) comboBoxGroupBy.SelectedItem = GroupingMode.None;
                            if (comboBoxThenBy.SelectedItem == null) comboBoxThenBy.SelectedItem = GroupingMode.None;

                            if ((GroupingMode)comboBoxThenBy.SelectedItem != GroupingMode.None || (GroupingMode)comboBoxGroupBy.SelectedItem != GroupingMode.None)
                            {




                                if ((GroupingMode)comboBoxThenBy.SelectedItem == GroupingMode.None && (GroupingMode)comboBoxGroupBy.SelectedItem != GroupingMode.None)
                                {
                                    GroupingMode mode = (GroupingMode)comboBoxGroupBy.SelectedItem;
                                    GroupingFunctions.GroupClashes(clashTestOriginal, mode, GroupingMode.None, (bool)keepExistingGroupsCheckBox.IsChecked, clashStatus);
                                }
                                else if ((GroupingMode)comboBoxGroupBy.SelectedItem == GroupingMode.None && (GroupingMode)comboBoxThenBy.SelectedItem != GroupingMode.None)
                                {
                                    GroupingMode mode = (GroupingMode)comboBoxThenBy.SelectedItem;
                                    GroupingFunctions.GroupClashes(clashTestOriginal, mode, GroupingMode.None, (bool)keepExistingGroupsCheckBox.IsChecked, clashStatus);
                                }
                                else
                                {
                                    GroupingMode byMode = (GroupingMode)comboBoxGroupBy.SelectedItem;
                                    GroupingMode thenByMode = (GroupingMode)comboBoxThenBy.SelectedItem;
                                    GroupingFunctions.GroupClashes(clashTestOriginal, byMode, thenByMode, (bool)keepExistingGroupsCheckBox.IsChecked, clashStatus);
                                }
                            }
                        }
                    }
                    catch { }



                }
                RegisterChanges();

                this.Close();
            }          
        }










        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            Keyboard.Focus(SearchText);
        }

        private void ClashGroupsForm_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close();
            }
        }

        /// <summary>
        /// Фильтрация по имени
        /// </summary>
        private void SearchText_Changed(object sender, RoutedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            string _searchName = textBox.Text.ToLower();

            ObservableCollection<CustomClashTest> filteredElement = new ObservableCollection<CustomClashTest>();

            foreach (CustomClashTest elemnt in ClashTests)
            {
                if (elemnt.ClashTest.DisplayName.ToLower().Contains(_searchName))
                {
                    filteredElement.Add(elemnt);
                }
            }

            ClashTestListBox.ItemsSource = filteredElement;
        }

        private void Ungroup_Button_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (ClashTestListBox.SelectedItems.Count != 0)
            {
                foreach (object selectedItem in ClashTestListBox.SelectedItems)
                {
                    CustomClashTest selectedClashTest = (CustomClashTest)selectedItem;
                    ClashTest clashTest = selectedClashTest.ClashTest;

                    try
                    {
                        if (clashTest.Children.Count != 0)
                        {
                            GroupingFunctions.UnGroupClashes(clashTest);
                        }
                    }
                    catch { }
                }
                RegisterChanges();

                this.Close();
            }
        }
        
        private void RegisterChanges()
        {
            GetClashTests();
            CheckPlugin();
            LoadComboBox();
        }
        
        private void CheckPlugin()
        {
            if (Application.MainDocument == null
                || Application.MainDocument.IsClear
                || Application.MainDocument.GetClash() == null
                || Application.MainDocument.GetClash().TestsData.Tests.Count == 0)
            {
                Group_Button.IsEnabled = false;
                comboBoxGroupBy.IsEnabled = false;
                comboBoxThenBy.IsEnabled = false;
                Ungroup_Button.IsEnabled = false;
            }
            else
            {
                Group_Button.IsEnabled = true;
                comboBoxGroupBy.IsEnabled = true;
                comboBoxThenBy.IsEnabled = true;
                Ungroup_Button.IsEnabled = true;
            }
        }

        private void LoadComboBox()
        {
            GroupByList.Clear();
            GroupThenList.Clear();
            foreach (GroupingMode mode in Enum.GetValues(typeof(GroupingMode)).Cast<GroupingMode>())
            {
                GroupThenList.Add(mode);
                GroupByList.Add(mode);
            }
            if (Application.MainDocument.Grids.ActiveSystem == null)
            {
                GroupByList.Remove(GroupingMode.GridIntersection);
                GroupByList.Remove(GroupingMode.Level);
                GroupThenList.Remove(GroupingMode.GridIntersection);
                GroupThenList.Remove(GroupingMode.Level);
            }
            comboBoxGroupBy.SelectedIndex = 0;
            comboBoxThenBy.SelectedIndex = 0;
        }
    }
}
