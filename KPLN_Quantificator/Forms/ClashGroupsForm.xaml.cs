using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api;
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
using Application = Autodesk.Navisworks.Api.Application;
using KPLN_Quantificator.Common;
using static KPLN_Quantificator.Common.Collections;

namespace KPLN_Quantificator.Forms
{
    /// <summary>
    /// Логика взаимодействия для ClashGroupsForm.xaml
    /// </summary>
    public partial class ClashGroupsForm : Window
    {
        public ObservableCollection<GroupingMode> GroupByList { get; set; } = new ObservableCollection<GroupingMode>();
        
        public ObservableCollection<GroupingMode> GroupThenList { get; set; } = new ObservableCollection<GroupingMode>();
        
        public ObservableCollection<CustomClashTest> ClashTests { get; set; } = new ObservableCollection<CustomClashTest>();
        //public ClashTest SelectedClashTest { get; set; }
        
        public ClashGroupsForm()
        {
            InitializeComponent();
            RegisterChanges();
            DataContext = this;
        }
        
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
            if (ClashTestListBox.SelectedItems.Count != 0)
            {
                foreach (object selectedItem in ClashTestListBox.SelectedItems)
                {
                    CustomClashTest selectedClashTest = (CustomClashTest)selectedItem;
                    ClashTest clashTest = selectedClashTest.ClashTest;
                    if (clashTest.Children.Count != 0)
                    {
                        if (comboBoxGroupBy.SelectedItem == null) comboBoxGroupBy.SelectedItem = GroupingMode.None;
                        if (comboBoxThenBy.SelectedItem == null) comboBoxThenBy.SelectedItem = GroupingMode.None;
                        if ((GroupingMode)comboBoxThenBy.SelectedItem != GroupingMode.None || (GroupingMode)comboBoxGroupBy.SelectedItem != GroupingMode.None)
                        {

                            if ((GroupingMode)comboBoxThenBy.SelectedItem == GroupingMode.None && (GroupingMode)comboBoxGroupBy.SelectedItem != GroupingMode.None)
                            {
                                GroupingMode mode = (GroupingMode)comboBoxGroupBy.SelectedItem;
                                GroupingFunctions.GroupClashes(clashTest, mode, GroupingMode.None, (bool)keepExistingGroupsCheckBox.IsChecked);
                            }
                            else if ((GroupingMode)comboBoxGroupBy.SelectedItem == GroupingMode.None && (GroupingMode)comboBoxThenBy.SelectedItem != GroupingMode.None)
                            {
                                GroupingMode mode = (GroupingMode)comboBoxThenBy.SelectedItem;
                                GroupingFunctions.GroupClashes(clashTest, mode, GroupingMode.None, (bool)keepExistingGroupsCheckBox.IsChecked);
                            }
                            else
                            {
                                GroupingMode byMode = (GroupingMode)comboBoxGroupBy.SelectedItem;
                                GroupingMode thenByMode = (GroupingMode)comboBoxThenBy.SelectedItem;
                                GroupingFunctions.GroupClashes(clashTest, byMode, thenByMode, (bool)keepExistingGroupsCheckBox.IsChecked);
                            }
                        }
                    }
                }
                RegisterChanges();
            }
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            Keyboard.Focus(SearchText);
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

                    if (clashTest.Children.Count != 0)
                    {
                        GroupingFunctions.UnGroupClashes(clashTest);
                    }
                }
                RegisterChanges();
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
