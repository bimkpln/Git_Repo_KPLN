using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;
using KPLN_Finishing.CommandTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using static KPLN_Finishing.Forms.MetaStack;

namespace KPLN_Finishing.Forms
{
    /// <summary>
    /// Логика взаимодействия для UpdateSetup.xaml
    /// </summary>
    public partial class UpdateSetup : Window
    {
        public bool HasErrors = false;
        public FilterRule Filter = new FilterRule(StorageType.None, null, "", true);
        public List<MetaRoom> GlobalRooms = new List<MetaRoom>();
        public UpdateSetup(List<WPFParameter> parameters, List<MetaRoom> rooms)
        {
            foreach (MetaRoom room in rooms)
            {
                GlobalRooms.Add(room);
            }
            InitializeComponent();
            this.cbxParameters_00.ItemsSource = parameters;
            this.cbxParameters_01.ItemsSource = parameters;
            this.cbxParameters_02.ItemsSource = parameters;
            this.cbxParameters_03.ItemsSource = new string[] { "Имена помещений", "Имена помещений + Номера", "Номера помещений + Имена", "Номера помещений" };
            this.cbxParameters_04.ItemsSource = new string[] { "1", "1.05", "1.10", "1.15", "1.20", "1.25" };
            this.cbxParameters_05.ItemsSource = parameters;
            this.cbxParameters_06.ItemsSource = parameters;
            this.cbxParameters_00.SelectedIndex = 0;
            this.cbxParameters_01.SelectedIndex = 0;
            this.cbxParameters_02.SelectedIndex = 0;
            this.cbxParameters_03.SelectedIndex = 0;
            this.cbxParameters_04.SelectedIndex = 0;
            this.cbxParameters_06.SelectedIndex = 0;
            this.cbxParameters_05_condition.SelectedIndex = 0;
        }
        private void OnClick(object sender, RoutedEventArgs args)
        {
            try
            {
                UpdatePeview();
            }
            catch (Exception e)
            {
                PrintError(e);
            }
        }
        private bool IsChecked(System.Windows.Controls.CheckBox box)
        {
            if ((bool)(box.IsChecked) == true) { return true; }
            return false;
        }
        private void OnSelectionChanged(object sender, SelectionChangedEventArgs args)
        {
            try
            {
                UpdatePeview();
            }
            catch (Exception e)
            {
                PrintError(e);
            }
        }
        private void UpdatePeview()
        {
            Preferences.IdlingCollections.Clear();
            HasErrors = false;
            List<MetaRoom> localRooms = new List<MetaRoom>();
            foreach (MetaRoom room in this.GlobalRooms)
            {
                if (Filter.PassesFilter(room.Room as Room))
                {
                    MetaRoom copy = room.GetCopy();
                    try
                    {
                        if ((this.cbxParameters_06.SelectedItem as WPFParameter).StorageType != StorageType.None)
                        {
                            copy.RecalculateNumber(this.cbxParameters_06.SelectedItem as WPFParameter);
                        }
                    }
                    catch (Exception) { }
                    localRooms.Add(copy);
                }
            }
            this.spExample.Children.Clear();
            if (!IsChecked(this.chbxWalls) && !IsChecked(this.chbxFloors) && !IsChecked(this.chbxCeilings) && !IsChecked(this.chbxPlinths))
            {
                return;
            }
            #region headers
            int wallColumn = 0, floorColumn = 0, ceilingColumn = 0, wallAreaColumn = 0, floorAreaColumn = 0, ceilingAreaColumn = 0, plinthColumn = 0, plinthAreaColumn = 0;
            int columns = 1;
            System.Windows.Controls.Grid grid = new System.Windows.Controls.Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition());
            if (IsChecked(this.chbxWalls))
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition());
                wallColumn = columns++;
                grid.ColumnDefinitions.Add(new ColumnDefinition());
                wallAreaColumn = columns++;
            }
            if (IsChecked(this.chbxFloors))
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition());
                floorColumn = columns++;
                grid.ColumnDefinitions.Add(new ColumnDefinition());
                floorAreaColumn = columns++;
            }
            if (IsChecked(this.chbxCeilings))
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition());
                ceilingColumn = columns++;
                grid.ColumnDefinitions.Add(new ColumnDefinition());
                ceilingAreaColumn = columns++;
            }
            if (IsChecked(this.chbxPlinths))
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition());
                plinthColumn = columns++;
                grid.ColumnDefinitions.Add(new ColumnDefinition());
                plinthAreaColumn = columns++;
            }
            System.Windows.Shapes.Rectangle rectangle = new System.Windows.Shapes.Rectangle();
            rectangle.Fill = Brushes.LightGray;
            rectangle.Margin = new Thickness() { Right = -1, Bottom = -1, Left = -1, Top = -1 };
            rectangle.RadiusX = 5;
            rectangle.RadiusY = 5;
            System.Windows.Controls.Grid.SetColumn(rectangle, 0);
            System.Windows.Controls.Grid.SetColumnSpan(rectangle, columns);
            grid.Children.Add(rectangle);
            TextBlock nameBlock = ExampleRows.GetTextBlock_Header("О_Группа");
            System.Windows.Controls.Grid.SetColumn(nameBlock, 0);
            grid.Children.Add(nameBlock);
            if (IsChecked(this.chbxWalls))
            {
                TextBlock wallBlock = ExampleRows.GetTextBlock_Header("О_ГОСТ_Описание стен");
                System.Windows.Controls.Grid.SetColumn(wallBlock, wallColumn);
                grid.Children.Add(wallBlock);
                TextBlock wallAreaBox = ExampleRows.GetTextBlock_Header("О_ГОСТ_Площадь стен");
                System.Windows.Controls.Grid.SetColumn(wallAreaBox, wallAreaColumn);
                grid.Children.Add(wallAreaBox);
            }
            if (IsChecked(this.chbxFloors))
            {
                TextBlock floorBlock = ExampleRows.GetTextBlock_Header("О_ГОСТ_Описание полов");
                System.Windows.Controls.Grid.SetColumn(floorBlock, floorColumn);
                grid.Children.Add(floorBlock);
                TextBlock floorAreaBlock = ExampleRows.GetTextBlock_Header("О_ГОСТ_Площадь полов");
                System.Windows.Controls.Grid.SetColumn(floorAreaBlock, floorAreaColumn);
                grid.Children.Add(floorAreaBlock);
            }
            if (IsChecked(this.chbxCeilings))
            {
                TextBlock ceilingsBlock = ExampleRows.GetTextBlock_Header("О_ГОСТ_Описание потолков");
                System.Windows.Controls.Grid.SetColumn(ceilingsBlock, ceilingColumn);
                grid.Children.Add(ceilingsBlock);
                TextBlock ceilingsAreaBlock = ExampleRows.GetTextBlock_Header("О_ГОСТ_Площадь потолков");
                System.Windows.Controls.Grid.SetColumn(ceilingsAreaBlock, ceilingAreaColumn);
                grid.Children.Add(ceilingsAreaBlock);
            }
            if (IsChecked(this.chbxPlinths))
            {
                TextBlock plinthBlock = ExampleRows.GetTextBlock_Header("О_ГОСТ_Описание плинтусов");
                System.Windows.Controls.Grid.SetColumn(plinthBlock, plinthColumn);
                grid.Children.Add(plinthBlock);
                TextBlock plinthAreaBlock = ExampleRows.GetTextBlock_Header("О_ГОСТ_Длина плинтусов");
                System.Windows.Controls.Grid.SetColumn(plinthAreaBlock, plinthAreaColumn);
                grid.Children.Add(plinthAreaBlock);
            }
            this.spExample.Children.Add(grid);
            this.spExample.Children.Add(new Separator());
            #endregion
            List<WPFParameter> parameters = new List<WPFParameter>();
            if (this.cbxParameters_00.SelectedIndex >= 1) { parameters.Add(this.cbxParameters_00.SelectedItem as WPFParameter); }
            if (this.cbxParameters_01.SelectedIndex >= 1) { parameters.Add(this.cbxParameters_01.SelectedItem as WPFParameter); }
            if (this.cbxParameters_02.SelectedIndex >= 1) { parameters.Add(this.cbxParameters_02.SelectedItem as WPFParameter); }
            double q = 1;
            switch (this.cbxParameters_04.SelectedIndex)
            {
                case 0:
                    q = 1;
                    break;
                case 1:
                    q = 1.05;
                    break;
                case 2:
                    q = 1.1;
                    break;
                case 3:
                    q = 1.15;
                    break;
                case 4:
                    q = 1.2;
                    break;
                case 5:
                    q = 1.25;
                    break;
                default:
                    q = 1;
                    break;
            }
            List<MetaRoomsStack> roomStacks = new List<MetaRoomsStack>();

            switch (parameters.Count)
            {
                case 0:
                    foreach (System.Windows.Controls.Grid g in ExampleRows.GetUpdatedUI(localRooms, parameters, IsChecked(this.chbxWalls), IsChecked(this.chbxFloors), IsChecked(this.chbxCeilings), IsChecked(this.chbxPlinths), this.cbxParameters_03.SelectedIndex, IsChecked(this.chbxUniqTypes), IsChecked(this.chbxCalculateResults), q, this))
                    {
                        this.spExample.Children.Add(g);
                        this.spExample.Children.Add(new Separator());
                    }
                    break;
                default:
                    foreach (string value in GetValues(parameters, localRooms))
                    {
                        foreach (MetaRoom room in localRooms)
                        {
                            List<string> values = new List<string>();
                            foreach (WPFParameter parameter in parameters)
                            { values.Add(parameter.GetStringValue(room.Room)); }
                            if (value == string.Join(" | ", values))
                            {
                                MetaRoomsStack foundRoomStack = MetaRoomsStack.GetStackByRoom(roomStacks, value);
                                if (foundRoomStack == null)
                                {
                                    roomStacks.Add(new MetaRoomsStack(room, value));
                                }
                                else
                                {
                                    foundRoomStack.AddRoom(room);
                                }
                            }
                        }
                    }
                    foreach (MetaRoomsStack stack in roomStacks)
                    {
                        this.spExample.Children.Add(ExampleRows.GetTextBlock_Header(stack.Key));
                        foreach (System.Windows.Controls.Grid g in ExampleRows.GetUpdatedUI(stack.Rooms, parameters, IsChecked(this.chbxWalls), IsChecked(this.chbxFloors), IsChecked(this.chbxCeilings), IsChecked(this.chbxPlinths), this.cbxParameters_03.SelectedIndex, IsChecked(this.chbxUniqTypes), IsChecked(this.chbxCalculateResults), q, this))
                        {
                            this.spExample.Children.Add(g);
                            this.spExample.Children.Add(new Separator());
                        }
                    }
                    break;
            }
        }
        private List<string> GetValues(List<WPFParameter> parameters, List<MetaRoom> rooms)
        {
            List<string> uniqValues = new List<string>();
            foreach (MetaRoom room in rooms)
            {
                List<string> values = new List<string>();
                foreach (WPFParameter parameter in parameters)
                {
                    values.Add(parameter.GetStringValue(room.Room));
                }
                string value = string.Join(" | ", values);
                if (!uniqValues.Contains(value))
                {
                    uniqValues.Add(value);
                }
            }
            uniqValues.Sort();
            return uniqValues;
        }
        private void OnFilterChange(object sender, SelectionChangedEventArgs e)
        {
            System.Windows.Controls.ComboBox senderBox = sender as System.Windows.Controls.ComboBox;
            if (senderBox == this.cbxParameters_05)
            {
                this.cbxParameters_05_value.Items.Clear();
                List<string> values = new List<string>();
                WPFParameter parameter = this.cbxParameters_05.SelectedItem as WPFParameter;
                foreach (MetaRoom room in this.GlobalRooms)
                {
                    string value = parameter.GetStringValue(room.Room);
                    if (!values.Contains(value))
                    { values.Add(value); }
                }
                values.Sort();
                foreach (string value in values)
                { this.cbxParameters_05_value.Items.Add(value); }
            }
        }
        private void OnFilterApply(object sender, RoutedEventArgs args)
        {
            WPFParameter pickedParam = this.cbxParameters_05.SelectedItem as WPFParameter;
            bool eqType = false;
            if (this.cbxParameters_05_condition.SelectedIndex == 0)
            { eqType = true; }
            try
            { Filter = new FilterRule(pickedParam.StorageType, pickedParam, this.cbxParameters_05_value.Text, eqType); }
            catch (Exception) { new FilterRule(StorageType.None, null, "", true); }
            try
            { UpdatePeview(); }
            catch (Exception e)
            { PrintError(e); }
        }
        private void OnOk(object sender, RoutedEventArgs e)
        {
            if (HasErrors)
            {
                this.Hide();
                TaskDialog td = new TaskDialog("Предупреждение");
                td.TitleAutoPrefix = false;
                td.MainContent = "Не все параметры заполнены!\nПодсказка: необходимо проверить все обозначенные красным цветом сроки.";
                td.FooterText = Names.task_dialog_hint;
                td.CommonButtons = TaskDialogCommonButtons.Ok;
                td.VerificationText = "Запустить в любом случае";
                td.Show();
                if (td.WasVerificationChecked())
                {
                    foreach (WritableCollection collection in Preferences.IdlingCollections)
                    {
                        Preferences.IdlingCollectionsRun.Add(collection);
                    }
                    this.Close();
                    return;
                }
                this.Show();
            }
            else
            {
                this.Hide();
                foreach (WritableCollection collection in Preferences.IdlingCollections)
                {
                    Preferences.IdlingCollectionsRun.Add(collection);
                }
                this.Close();
            }
        }
    }
    public class MetaRoomsStack
    {
        public readonly string Key = "";
        public readonly List<MetaRoom> Rooms = new List<MetaRoom>();
        public MetaRoomsStack(MetaRoom room, string key)
        {
            Rooms.Add(room);
            Key = key;
        }
        public static MetaRoomsStack GetStackByRoom(List<MetaRoomsStack> stacks, string roomKey)
        {
            foreach (MetaRoomsStack stack in stacks)
            {
                if (stack.Key == roomKey)
                {
                    return stack;
                }
            }
            return null;
        }
        public void AddRoom(MetaRoom room)
        {
            Rooms.Add(room);
        }
    }
    public class FilterRule
    {
        public StorageType Type = StorageType.None;
        private string Value = "";
        private bool EQ = false;
        private WPFParameter Parameter = new WPFParameter();
        public bool PassesFilter(Room room)
        {
            if (Type == StorageType.None)
            { return true; }
            else
            {
                if (EQ)
                {
                    if (Parameter.GetStringValue(room) == Value)
                    {
                        return true;
                    }
                    else { return false; }
                }
                else
                {
                    if (Parameter.GetStringValue(room) != Value)
                    {
                        return true;
                    }
                    else { return false; }
                }
            }
        }
        public FilterRule(StorageType type, WPFParameter parameter, string value, bool ifTrue)
        {
            Type = type;
            Parameter = parameter;
            Value = value;
            EQ = ifTrue;
        }
    }
    public static class ExampleRows
    {
        private static int size = 12;
        public static TextBlock GetTextBlock(string text)
        {
            TextBlock tb = new TextBlock();
            tb.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
            tb.TextWrapping = TextWrapping.Wrap;
            tb.Text = text;
            tb.FontSize = size;
            Thickness margin = new Thickness();
            margin.Left = 5;
            margin.Right = 5;
            tb.Margin = margin;
            return tb;
        }
        public static TextBlock GetTextBlock_Header(string text)
        {
            TextBlock tb = new TextBlock();
            tb.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
            tb.TextWrapping = TextWrapping.Wrap;
            tb.Text = text;
            tb.FontSize = 12;
            tb.FontWeight = FontWeights.Bold;
            Thickness margin = new Thickness();
            margin.Left = 5;
            margin.Right = 5;
            tb.Margin = margin;
            return tb;
        }
        public static MetaStack GetStackByRoom(List<MetaStack> stacks, MetaRoom room)
        {
            foreach (MetaStack stack in stacks)
            {
                if (stack.Key == room.Key)
                {
                    return stack;
                }
            }
            return null;
        }
        public static List<System.Windows.Controls.Grid> GetUpdatedUI(List<MetaRoom> rooms, List<WPFParameter> groupParameters, bool walls, bool floors, bool ceilings, bool plinth, int typeIndex, bool uniqTypeSets, bool calculateResults, double q, UpdateSetup sender)
        {
            string wallsDescription = null;
            string floorDescription = null;
            string ceilingDescription = null;
            string plinthsDescription = null;
            string wallsArea = null;
            string floorArea = null;
            string ceilingArea = null;
            string plinthsArea = null;
            string group = null;
            List<MetaStack> Stacks = new List<MetaStack>();
            foreach (MetaRoom room in rooms)
            {
                room.CalculateKey(groupParameters, uniqTypeSets, walls, floors, ceilings, plinth);
                MetaStack foundedStack = GetStackByRoom(Stacks, room);
                if (foundedStack != null)
                {
                    foundedStack.AddRoom(room);
                }
                else
                {
                    MetaStack stack = new MetaStack(room.Key);
                    stack.AddRoom(room);
                    Stacks.Add(stack);
                }
            }
            foreach (MetaStack stack in Stacks)
            {
                stack.Update();
            }
            int wallColumn = 0, floorColumn = 0, ceilingColumn = 0, wallAreaColumn = 0, floorAreaColumn = 0, ceilingAreaColumn = 0, plinthColumn = 0, plinthAreaColumn = 0;
            List<System.Windows.Controls.Grid> rows = new List<System.Windows.Controls.Grid>();
            foreach (MetaStack stack in Stacks)
            {
                size = 12;
                int columns = 1;
                System.Windows.Controls.Grid grid = new System.Windows.Controls.Grid();
                grid.ColumnDefinitions.Add(new ColumnDefinition());
                if (walls)
                {
                    size--;
                    grid.ColumnDefinitions.Add(new ColumnDefinition());
                    wallColumn = columns++;
                    grid.ColumnDefinitions.Add(new ColumnDefinition());
                    wallAreaColumn = columns++;
                }
                if (floors)
                {
                    size--;
                    grid.ColumnDefinitions.Add(new ColumnDefinition());
                    floorColumn = columns++;
                    grid.ColumnDefinitions.Add(new ColumnDefinition());
                    floorAreaColumn = columns++;
                }
                if (ceilings)
                {
                    size--;
                    grid.ColumnDefinitions.Add(new ColumnDefinition());
                    ceilingColumn = columns++;
                    grid.ColumnDefinitions.Add(new ColumnDefinition());
                    ceilingAreaColumn = columns++;
                }
                if (plinth)
                {
                    size--;
                    grid.ColumnDefinitions.Add(new ColumnDefinition());
                    plinthColumn = columns++;
                    grid.ColumnDefinitions.Add(new ColumnDefinition());
                    plinthAreaColumn = columns++;
                }
                TypeValue wallValues = stack.GetValue(FinishingType.Walls, calculateResults, q);
                TypeValue floorValues = stack.GetValue(FinishingType.Floors, calculateResults, q);
                TypeValue ceilingValues = stack.GetValue(FinishingType.Ceilings, calculateResults, q);
                TypeValue plinthValues = stack.GetValue(FinishingType.Plinths, calculateResults, q);
                if (walls)
                {
                    wallsDescription = wallValues.firstItem;
                    wallsArea = wallValues.lastItem;
                }
                if (floors)
                {
                    floorDescription = floorValues.firstItem;
                    floorArea = floorValues.lastItem;
                }
                if (ceilings)
                {
                    ceilingDescription = ceilingValues.firstItem;
                    ceilingArea = ceilingValues.lastItem;
                }
                if (plinth)
                {
                    plinthsDescription = plinthValues.firstItem;
                    plinthsArea = plinthValues.lastItem;
                }
                if ((walls && wallValues.lastItem.Contains("Не заполнено")) || (floorValues.lastItem.Contains("Не заполнено") && floors) || (ceilingValues.lastItem.Contains("Не заполнено") && ceilings) || (plinthValues.lastItem.Contains("Не заполнено") && plinth))
                {
                    sender.HasErrors = true;
                    System.Windows.Shapes.Rectangle rectangle = new System.Windows.Shapes.Rectangle();
                    rectangle.Fill = Brushes.LightPink;
                    rectangle.Margin = new Thickness() { Right = 2, Bottom = 2, Left = 2, Top = 2 };
                    rectangle.RadiusX = 5;
                    rectangle.RadiusY = 5;
                    System.Windows.Controls.Grid.SetColumn(rectangle, 0);
                    System.Windows.Controls.Grid.SetColumnSpan(rectangle, columns);
                    grid.Children.Add(rectangle);
                }
                group = stack.GetName(typeIndex);
                TextBlock nameBlock = GetTextBlock(stack.GetName(typeIndex));
                System.Windows.Controls.Grid.SetColumn(nameBlock, 0);
                grid.Children.Add(nameBlock);
                if (walls)
                {
                    TextBlock wallBlock = GetTextBlock(wallValues.firstItem);
                    System.Windows.Controls.Grid.SetColumn(wallBlock, wallColumn);
                    grid.Children.Add(wallBlock);
                    TextBlock wallAreaBox = GetTextBlock(wallValues.lastItem);
                    System.Windows.Controls.Grid.SetColumn(wallAreaBox, wallAreaColumn);
                    grid.Children.Add(wallAreaBox);
                }
                if (floors)
                {
                    TextBlock floorBlock = GetTextBlock(floorValues.firstItem);
                    System.Windows.Controls.Grid.SetColumn(floorBlock, floorColumn);
                    grid.Children.Add(floorBlock);
                    TextBlock floorAreaBlock = GetTextBlock(floorValues.lastItem);
                    System.Windows.Controls.Grid.SetColumn(floorAreaBlock, floorAreaColumn);
                    grid.Children.Add(floorAreaBlock);
                }
                if (ceilings)
                {
                    TextBlock ceilingsBlock = GetTextBlock(ceilingValues.firstItem);
                    System.Windows.Controls.Grid.SetColumn(ceilingsBlock, ceilingColumn);
                    grid.Children.Add(ceilingsBlock);
                    TextBlock ceilingsAreaBlock = GetTextBlock(ceilingValues.lastItem);
                    System.Windows.Controls.Grid.SetColumn(ceilingsAreaBlock, ceilingAreaColumn);
                    grid.Children.Add(ceilingsAreaBlock);
                }
                if (plinth)
                {
                    TextBlock plinthBlock = GetTextBlock(plinthValues.firstItem);
                    System.Windows.Controls.Grid.SetColumn(plinthBlock, plinthColumn);
                    grid.Children.Add(plinthBlock);
                    TextBlock plinthAreaBlock = GetTextBlock(plinthValues.lastItem);
                    System.Windows.Controls.Grid.SetColumn(plinthAreaBlock, plinthAreaColumn);
                    grid.Children.Add(plinthAreaBlock);
                }
                Preferences.IdlingCollections.Add(new WritableCollection(stack.Rooms,
                                                                    wallsDescription,
                                                                    wallsArea,
                                                                    floorDescription,
                                                                    floorArea,
                                                                    ceilingDescription,
                                                                    ceilingArea,
                                                                    plinthsDescription,
                                                                    plinthsArea,
                                                                    group));
                rows.Add(grid);
            }
            return rows;
        }
    }
    public class LocalMetaStack
    {
        public double Value = 0;
        public FinishingType Type = FinishingType.Null;
        public readonly string Key = "";
        public readonly string Mark = "Не заполнено";
        public readonly string Description = "Не заполнено";
        public void AddElement(MetaElement element)
        {
            switch (element.Type)
            {
                case FinishingType.Plinths:
                    Value += element.Length;
                    break;
                default:
                    Value += element.Area;
                    break;
            }
        }
        public LocalMetaStack(MetaElement element)
        {
            Key = element.Mark + element.TypeIdString;
            Mark = element.Mark;
            Description = element.Description;
            Type = element.Type;
            switch (element.Type)
            {
                case FinishingType.Plinths:
                    Value += element.Length;
                    break;
                default:
                    Value += element.Area;
                    break;
            }
        }
    }
    public class MetaStack
    {
        #region variables
        public readonly string Key;
        public List<string> Names = new List<string>();
        public List<string> Numbers = new List<string>();
        public List<MetaRoom> Rooms = new List<MetaRoom>();
        public List<MetaElement> Elements = new List<MetaElement>();
        public List<LocalMetaStack> WallStacks = new List<LocalMetaStack>();
        public List<LocalMetaStack> FloorStacks = new List<LocalMetaStack>();
        public List<LocalMetaStack> CeilingStacks = new List<LocalMetaStack>();
        public List<LocalMetaStack> PlinthStacks = new List<LocalMetaStack>();
        #endregion
        public MetaStack(string key)
        {
            Key = key;
        }
        public void AddRoom(MetaRoom room)
        {
            if (room.Key != Key) { throw new Exception("Ключ помещения не совпадает с ключем коллекции!"); }
            Rooms.Add(room);
            foreach (MetaElement element in room.Elements)
            {
                Elements.Add(element);
            }
            if (!Names.Contains(room.Name))
            {
                Names.Add(room.Name);
            }
            if (!Numbers.Contains(room.Number))
            {
                Numbers.Add(room.Number);
            }
        }
        public string GetName(int typeIndex)
        {
            switch (typeIndex)
            {
                case 0:
                    return string.Join(", ", Names);
                case 1:
                    return string.Join("\n\n", new string[] { string.Join(", ", Names), string.Join(", ", Numbers) });
                case 2:
                    return string.Join("\n\n", new string[] { string.Join(", ", Numbers), string.Join(", ", Names) });
                case 3:
                    return string.Join(", ", Numbers);
                default:
                    return "";
            }
        }
        public LocalMetaStack GetStackByElement(MetaElement element, List<LocalMetaStack> stacks)
        {
            foreach (LocalMetaStack stack in stacks)
            {
                if (stack.Key == element.Mark + element.TypeIdString)
                {
                    return stack;
                }
            }
            return null;
        }
        private void AddToLocalStack(List<LocalMetaStack> stacks, MetaElement element)
        {
            LocalMetaStack stack = GetStackByElement(element, stacks);
            if (stack != null)
            {
                stack.AddElement(element);
            }
            else
            {
                stacks.Add(new LocalMetaStack(element));
            }
            List<string> keys = new List<string>();
            foreach (LocalMetaStack localStack in stacks)
            { keys.Add(localStack.Key); }
            keys.Sort();
            List<LocalMetaStack> sortedStacks = new List<LocalMetaStack>();
            foreach (string key in keys)
            {
                foreach (LocalMetaStack localStack in stacks)
                {
                    if (localStack.Key == key) { sortedStacks.Add(localStack); }
                }
            }
            stacks.Clear();
            foreach (LocalMetaStack sortedStack in sortedStacks)
            { stacks.Add(sortedStack); }
        }
        public void Update()
        {
            WallStacks.Clear();
            FloorStacks.Clear();
            CeilingStacks.Clear();
            PlinthStacks.Clear();
            foreach (MetaElement element in Elements)
            {
                switch (element.Type)
                {
                    case FinishingType.Walls:
                        AddToLocalStack(WallStacks, element);
                        break;
                    case FinishingType.Floors:
                        AddToLocalStack(FloorStacks, element);
                        break;
                    case FinishingType.Ceilings:
                        AddToLocalStack(CeilingStacks, element);
                        break;
                    case FinishingType.Plinths:
                        AddToLocalStack(PlinthStacks, element);
                        break;
                    default:
                        break;
                }
            }
        }
        public class TypeValue
        {
            public string firstItem { get; set; }
            public string lastItem { get; set; }
            public TypeValue(string first, string second)
            {
                firstItem = first;
                lastItem = second;
            }
            public TypeValue()
            {
                firstItem = null;
                lastItem = null;
            }
            public void SetFirst(string value)
            {
                firstItem = value;
            }
            public void SetLast(string value)
            {
                lastItem = value;
            }
        }
        public TypeValue GetValue(FinishingType type, bool calculateResults, double q)
        {
            List<LocalMetaStack> stacks = new List<LocalMetaStack>();
            switch (type)
            {
                case FinishingType.Walls:
                    stacks = WallStacks;
                    break;
                case FinishingType.Floors:
                    stacks = FloorStacks;
                    break;
                case FinishingType.Ceilings:
                    stacks = CeilingStacks;
                    break;
                case FinishingType.Plinths:
                    stacks = PlinthStacks;
                    break;
                default:
                    break;
            }
            TypeValue result = new TypeValue();
            string description = "";
            string area = "";
            string units = "";
            double commonArea = 0;
            foreach (LocalMetaStack stack in stacks)
            {
                double div = 1;
                units = "";
                if (stack.Type == FinishingType.Plinths)
                {
                    units = "пог.м";
                    div = 0.3048;
                }
                else
                {
                    units = "м²";
                    div = 0.092903;
                }
                int n = stack.Description.Count(x => x == '\n');
                description += string.Format("Тип: «{0}»\n{1}\n\n", stack.Mark, stack.Description);
                area += string.Format("Тип: «{0}»\n{1} {2}\n\n", stack.Mark, Math.Round(Math.Round(stack.Value * div, 2) * q, 2).ToString(), units);
                commonArea += Math.Round(Math.Round(stack.Value * div, 2) * q, 2);
                while (n > 0)
                {
                    area += "\n";
                    n -= 1;
                }
            }
            if (calculateResults && stacks.Count > 1)
            {
                area += string.Format("Итого: {0} {1}", commonArea.ToString(), units);
            }
            result.SetFirst(description);
            result.SetLast(area);
            return result;
        }
    }
    public class WritableCollection : AbstractOnIdlingCommand
    {
        public static int Counter = 0;
        public static Guid Guid = Guid.NewGuid();
        private List<ElementId> Rooms = new List<ElementId>();
        private string Group { get; }
        private string RowId { get; }
        private string WallsDescription { get; }
        private string WallsArea { get; }
        private string FloorsDescription { get; }
        private string FloorsArea { get; }
        private string CeilingsDescription { get; }
        private string CeilingsArea { get; }
        private string PlinthsDescription { get; }
        private string PlinthsArea { get; }
        public static void ResetSession()
        {
            Counter = 0;
            Guid = Guid.NewGuid();
        }
        public WritableCollection(List<MetaRoom> rooms,
                                                string wallsDescription,
                                                string wallsArea,
                                                string floorsDescription,
                                                string floorsArea,
                                                string ceilingsDescription,
                                                string ceilingsArea,
                                                string plinthsDescription,
                                                string plinthsArea,
                                                string group)
        {
            RowId = string.Format("{0} : {1}", Guid.ToString(), Counter++.ToString());
            foreach (MetaRoom room in rooms)
            {
                Rooms.Add(room.Room.Id);
            }
            WallsDescription = wallsDescription;
            WallsArea = wallsArea;
            FloorsDescription = floorsDescription;
            FloorsArea = floorsArea;
            CeilingsDescription = ceilingsDescription;
            CeilingsArea = ceilingsArea;
            PlinthsDescription = plinthsDescription;
            PlinthsArea = plinthsArea;
            Group = group;

        }
        public override void Execute(UIApplication uiapp)
        {
            try
            {
                foreach (ElementId roomId in Rooms)
                {
                    Room room = uiapp.ActiveUIDocument.Document.GetElement(roomId) as Room;
                    try
                    {
                        room.LookupParameter("О_ПОМ_Ведомость").Set(RowId);
                        if (Group != null)
                        {
                            room.LookupParameter("О_ПОМ_Группа").Set(Group);
                        }
                        if (WallsDescription != null && WallsArea != null)
                        {
                            room.LookupParameter("О_ПОМ_ГОСТ_Описание стен").Set(WallsDescription);
                            room.LookupParameter("О_ПОМ_ГОСТ_Площадь стен_Текст").Set(WallsArea);
                        }
                        if (FloorsDescription != null && FloorsArea != null)
                        {
                            room.LookupParameter("О_ПОМ_ГОСТ_Описание полов").Set(FloorsDescription);
                            room.LookupParameter("О_ПОМ_ГОСТ_Площадь полов_Текст").Set(FloorsArea);
                        }
                        if (CeilingsDescription != null && CeilingsArea != null)
                        {
                            room.LookupParameter("О_ПОМ_ГОСТ_Описание потолков").Set(CeilingsDescription);
                            room.LookupParameter("О_ПОМ_ГОСТ_Площадь потолков_Текст").Set(CeilingsArea);
                        }
                        if (PlinthsDescription != null && PlinthsArea != null)
                        {
                            room.LookupParameter("О_ПОМ_ГОСТ_Описание плинтусов").Set(PlinthsDescription);
                            room.LookupParameter("О_ПОМ_ГОСТ_Длина плинтусов_Текст").Set(PlinthsArea);
                        }
                    }
                    catch (Exception e) { PrintError(e); }
                }
            }
            catch (Exception e) { PrintError(e); }
        }
    }
    public class MetaRoom
    {
        public int Id { get; set; }
        public Element Room { get; }
        public string Name { get; }
        public string Number { get; set; }
        public string Key { get { return InnerKey; } }
        private string InnerKey { get; set; }
        public List<MetaElement> Elements { get; }
        private List<string> UniqTypes { get; }
        private readonly List<string> WallTypes = new List<string>();
        private readonly List<string> FloorTypes = new List<string>();
        private readonly List<string> CeilingTypes = new List<string>();
        private readonly List<string> PlinthTypes = new List<string>();
        public string CalculateKey(List<WPFParameter> parameters, bool uniqTypeSets, bool walls, bool floors, bool ceilings, bool plinth)
        {
            if (parameters.Count == 0 && !uniqTypeSets)
            {
                return "";
            }
            string key = "";
            List<string> values = new List<string>();
            foreach (WPFParameter parameter in parameters)
            {
                values.Add(parameter.GetStringValue(Room));
            }
            if (uniqTypeSets)
            {
                if (walls)
                {
                    key += string.Join("_", values);
                    WallTypes.Sort();
                    key += string.Join("_", WallTypes);
                }
                if (floors)
                {
                    key += string.Join("_", values);
                    FloorTypes.Sort();
                    key += string.Join("_", FloorTypes);
                }
                if (ceilings)
                {
                    key += string.Join("_", values);
                    CeilingTypes.Sort();
                    key += string.Join("_", CeilingTypes);
                }
                if (plinth)
                {
                    key += string.Join("_", values);
                    PlinthTypes.Sort();
                    key += string.Join("_", PlinthTypes);
                }
            }
            InnerKey = key;
            return key;
        }
        public MetaRoom(Room room)
        {
            Id = room.Id.IntegerValue;
            UniqTypes = new List<string>();
            Elements = new List<MetaElement>();
            InnerKey = "";
            Room = room;
            Name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
            Number = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString();

        }
        public MetaRoom GetCopy()
        {
            MetaRoom copy = new MetaRoom(this.Room as Room);
            copy.InnerKey = InnerKey;
            foreach (MetaElement element in Elements)
            {
                copy.AddElement(element);
            }
            return copy;
        }
        public void RecalculateNumber(WPFParameter parameter)
        {
            Number = parameter.GetStringValue(Room);
        }
        public void AddElement(MetaElement element)
        {
            Elements.Add(element);
            if (!UniqTypes.Contains(element.TypeIdString))
            {
                UniqTypes.Add(element.TypeIdString);
                switch (element.Type)
                {
                    case FinishingType.Walls:
                        WallTypes.Add(element.TypeIdString);
                        break;
                    case FinishingType.Floors:
                        FloorTypes.Add(element.TypeIdString);
                        break;
                    case FinishingType.Ceilings:
                        CeilingTypes.Add(element.TypeIdString);
                        break;
                    case FinishingType.Plinths:
                        PlinthTypes.Add(element.TypeIdString);
                        break;
                    default:
                        break;
                }
            }
        }
    }
    public class MetaElement
    {
        public readonly string TypeIdString = "-1";
        public readonly double Area = 0.0;
        public readonly double Length = 0.0;
        public readonly string Description = "Не заполнено";
        public readonly string Mark = "Не заполнено";
        public readonly FinishingType Type = FinishingType.Null;
        private bool ContainsChat(string v)
        {
            foreach (char c in "0123456789.,abcdefghijklmnopqrstuvwxyzабвгдеёжзийклмнопрстуфхцчшщъыьэюя-=+")
            {
                if (v.ToLower().Contains(c))
                {
                    return true;
                }
            }
            return false;
        }
        public MetaElement(Element element)
        {
            TypeIdString = Tools.GetTypeElement(element).Id.ToString();
            Element type = Tools.GetTypeElement(element);
            try
            {
                string v = type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK).AsString();
                if (v != "" && ContainsChat(v))
                {
                    Mark = v;
                }

            }
            catch (Exception)
            {
                Mark = "Не заполнено";
            }

            try
            {
                string v = type.LookupParameter("О_Описание").AsString();
                if (v != "")
                {
                    Description = type.LookupParameter("О_Описание").AsString();
                }
            }
            catch (Exception)
            {
                Description = "Не заполнено";
            }

            switch (element.Category.Id.IntegerValue)
            {
                case -2000011://Walls
                    try
                    {
                        if (type.LookupParameter("О_Плинтус").AsInteger() == 1 && type.LookupParameter("О_Плинтус_Высота").AsDouble() > 0)
                        {
                            Type = FinishingType.Plinths;
                            Length = (element).get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() / type.LookupParameter("О_Плинтус_Высота").AsDouble();
                        }
                        else
                        {
                            Type = FinishingType.Walls;
                            Area = (element).get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                        }
                    }
                    catch (Exception)
                    {
                        Type = FinishingType.Walls;
                        Area = element.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                    }
                    break;
                case -2000032://Floors
                    Type = FinishingType.Floors;
                    Area = element.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                    break;
                case -2000038://Ceilings
                    Type = FinishingType.Ceilings;
                    Area = element.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                    break;
                default:
                    break;
            }

        }
    }
    public enum FinishingType { Walls, Floors, Ceilings, Plinths, Null }
}
