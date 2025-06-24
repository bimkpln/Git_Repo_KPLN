using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Text.RegularExpressions;
using System.Globalization;

using Grid = System.Windows.Controls.Grid;
using TextBox = System.Windows.Controls.TextBox;


namespace KPLN_BIMTools_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для batchAddingParametersWindowCopyFromFamily.xaml
    /// </summary>
    public partial class batchAddingParametersWindowCopyFromFamily : Window
    {
        private readonly UIApplication _uiApp;

        public batchAddingParametersWindowCopyFromFamily(UIApplication uiApp)
        {
            InitializeComponent();

            _uiApp = uiApp;

            LoadFamilyDocuments();
        }

        // Загрузка имён всех существующих семейств
        private void LoadFamilyDocuments()
        {
            var familyDocs = _uiApp.Application.Documents
                .Cast<Document>()
                .Where(doc => doc.IsFamilyDocument)
                .Select(doc => doc.Title)
                .Distinct()
                .ToList();

            CB_familyName.ItemsSource = familyDocs;
            Document activeDoc = _uiApp.ActiveUIDocument?.Document;
            if (activeDoc != null && activeDoc.IsFamilyDocument)
            {
                CB_familyName.SelectedItem = activeDoc.Title;
                LoadFamilyParameters(activeDoc);
            }
            else
            {
                ParameterListPanel.Children.Clear();
            }
        }

        // Загрузка параметров из семейства
        private void LoadFamilyParameters(Document familyDoc)
        {
            if (!familyDoc.IsFamilyDocument || familyDoc.FamilyManager == null)
                return;

            FamilyManager famMgr = familyDoc.FamilyManager;
            FamilyType currentType = famMgr.CurrentType;

            if (currentType == null)
                return;

            ParameterListPanel.Children.Clear();

            foreach (FamilyParameter param in famMgr.Parameters.Cast<FamilyParameter>().OrderBy(p => p.Definition.Name))
            {
                var row = new Grid
                {

                    Margin = new Thickness(0, 0, 0, 5)
                };
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(470) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(195) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });

                var isInstance = param.IsInstance;
                var isShared = param.IsShared;

                Brush colorString = IsSharedParameterExistsInFile(_uiApp, param);
             
#if (Debug2020 || Revit2020)
                var group = param.Definition.ParameterGroup; 
var typeInfo = param.Definition.ParameterType.ToString();
#endif

#if (Debug2023 || Revit2023)
                var group = param.Definition.GetGroupTypeId(); 
                var typeInfo = param.Definition.GetDataType().TypeId;
#endif

                var paramName = new TextBlock
                {
                    Text = param.Definition.Name,
                    Tag = new
                    {                       
                        Param = param,
                        IsShared = isShared,
                        Group = group,
                        IsInstance = isInstance
                    },

#if (Debug2020 || Revit2020)
                    ToolTip = $"Тип данных: {param.Definition.ParameterType}",
#endif
#if (Debug2023 || Revit2023)
                    ToolTip = $"Тип данных: {param.Definition.GetDataType().TypeId}",
#endif

                    FontSize = 12,
                    Foreground = colorString,
                    Padding = new Thickness(10, 0, 0, 0),
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetColumn(paramName, 0);
                row.Children.Add(paramName);

                var valueBox = new System.Windows.Controls.TextBox
                {
                    Text = GetFamilyParamValue(_uiApp.ActiveUIDocument?.Document, currentType, param),
#if (Debug2020 || Revit2020)
                    Tag = param.Definition.ParameterType,
#endif
#if (Debug2023 || Revit2023)
                    Tag = param.Definition.GetDataType().TypeId,
#endif
                    FontSize = 12,
                    Margin = new Thickness(2, 0, 2, 0),
                    Background = colorString,
                    IsEnabled = false                  
            }
            ;
                Grid.SetColumn(valueBox, 1);
                row.Children.Add(valueBox);

                var cbValue = new CheckBox
                {
                    HorizontalAlignment = HorizontalAlignment.Center,
                    IsEnabled = false
                };

                var cbParam = new CheckBox
                {
                    HorizontalAlignment = HorizontalAlignment.Center,
                    IsEnabled = (colorString == Brushes.Red) ? false : true,
                };

                var tbValue = new TextBlock
                {
                    Text = "Не копируется",
                    FontSize = 10,
                    Foreground = colorString,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Padding = new Thickness(0, 2, 0, 0)
                };
                var tbParam = new TextBlock
                {
                    Text = "Не копируется",
                    FontSize = 10,
                    Foreground = colorString,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Padding = new Thickness(0, 2, 0, 0)
                };
              
                Grid.SetColumn(cbParam, 3);


                if (colorString == Brushes.Red)
                {
                    Grid.SetColumn(tbValue, 2);
                    Grid.SetColumn(tbParam, 3);
                    row.Children.Add(tbValue);
                    row.Children.Add(tbParam);
                }
                else if (colorString == Brushes.Yellow)
                {
                    Grid.SetColumn(tbValue, 2);
                    Grid.SetColumn(cbParam, 3);
                    row.Children.Add(tbValue);
                    row.Children.Add(cbParam);
                }
                else
                {
                    Grid.SetColumn(cbValue, 2);
                    Grid.SetColumn(cbParam, 3);
                    row.Children.Add(cbValue);
                    row.Children.Add(cbParam);
                }

                cbParam.Checked += (s, e) =>
                {
                    cbValue.IsEnabled = true;
                };

                cbParam.Unchecked += (s, e) =>
                {
                    cbValue.IsChecked = false;
                    cbValue.IsEnabled = false;
                    valueBox.IsEnabled = false;
                };

                cbValue.Checked += (s, e) =>
                {
                    valueBox.IsEnabled = true;
                };

                cbValue.Unchecked += (s, e) =>
                {
                    valueBox.IsEnabled = false;
                };

                ParameterListPanel.Children.Add(row);
            }
        }

        // Проверка на параметр из текущего ФОП
        Brush IsSharedParameterExistsInFile(UIApplication uiApp, FamilyParameter param)
        {
            try
            {
                DefinitionFile sharedParameterFile = uiApp.Application.OpenSharedParameterFile();
                if (param.IsShared && sharedParameterFile == null)
                    return Brushes.Red;

#if (Debug2020 || Revit2020)
                if (param.Definition.ParameterType == ParameterType.Image ||
                    param.Definition.ParameterType == ParameterType.LoadClassification)
                {
                    return Brushes.Yellow;
                }
#endif
#if (Debug2023 || Revit2023)
                if (param.Definition.GetDataType() == SpecTypeId.Reference.Image ||
                    param.Definition.GetDataType() == SpecTypeId.Reference.LoadClassification)
                {
                    return Brushes.Yellow;
                }
#endif

                if (param.IsShared)
                {
                    foreach (DefinitionGroup groupDef in sharedParameterFile.Groups)
                    {
                        foreach (Definition def in groupDef.Definitions)
                        {
                            if (def is ExternalDefinition extDef
                                && extDef.GUID == param.GUID
                                && extDef.Name == param.Definition.Name
#if (Debug2020 || Revit2020)
                        && extDef.ParameterType == param.Definition.ParameterType
#endif
#if (Debug2023 || Revit2023)
                        && extDef.GetDataType().TypeId == param.Definition.GetDataType().TypeId
#endif
                                )
                            {
                                return Brushes.LightBlue;
                            }
                        }
                    }
                }
                else
                {
                    return Brushes.White;
                }

                return Brushes.Red;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", $"{ex.Message}");
                return Brushes.Red;
            }
        }

        // Чтение значений параметров
        private string GetFamilyParamValue(Document doc, FamilyType type, FamilyParameter param)
        {
            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String:
                        return type.AsString(param) ?? "";

                    case StorageType.Double:
                        double? maybeVal = type.AsDouble(param);
                        if (!maybeVal.HasValue)
                        {

                            return "";
                        }

                        double val = maybeVal.Value;


#if (Debug2020 || Revit2020)
                        UnitType unitType = param.Definition.UnitType;
                        string formatted = UnitFormatUtils.Format(doc.GetUnits(), unitType, val, false, false);
                        return Regex.Match(formatted, @"[\d\.,\-]+").Value;
#endif
#if (Debug2023 || Revit2023)
                        ForgeTypeId spec = param.Definition.GetDataType();
                        if (UnitUtils.IsMeasurableSpec(spec))
                        {
                            string formatted = UnitFormatUtils.Format(doc.GetUnits(), spec, val, forEditing: false);
                            return Regex.Match(formatted, @"[\d\.,\-]+").Value;
                        }
                        return val.ToString(CultureInfo.InvariantCulture);
#endif


                    case StorageType.Integer:
                        return type.AsInteger(param).ToString();

                    case StorageType.ElementId:
                        return type.AsElementId(param).IntegerValue.ToString();

                    default:
                        return "";
                }
            }
            catch
            {
                return "[Ошибка]";
            }
        }

        //// XAML. Обновление параметров семейства
        private void CB_familyName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string selectedTitle = CB_familyName.SelectedItem as string;

            if (string.IsNullOrWhiteSpace(selectedTitle))
                return;

            Document selectedDoc = _uiApp.Application.Documents
                .Cast<Document>()
                .FirstOrDefault(d => d.IsFamilyDocument && d.Title == selectedTitle);

            if (selectedDoc != null)
            {
                LoadFamilyParameters(selectedDoc);
            }
        }

        //// XAML. Открытие следующего диалогового окна
        private void ButtonNextStep_Click(object sender, RoutedEventArgs e)
        {
            var selectedParameters = new Dictionary<FamilyParameter, (bool isShared, Guid guid, object Group, bool? CopyValue, object ParamType, bool isInstance, string Value)>();

            foreach (Grid row in ParameterListPanel.Children.OfType<Grid>())
            {
                var paramName = row.Children[0] as TextBlock;
                if (paramName?.Tag == null) continue;

                dynamic tag = paramName.Tag;
                FamilyParameter param = tag.Param;
                bool isShared = tag.IsShared;
                bool isInstance = tag.IsInstance;

#if (Debug2020 || Revit2020)
                var group = (BuiltInParameterGroup)tag.Group;
                var paramType = param.Definition.ParameterType; 
#endif
#if (Debug2023 || Revit2023)
                var group = (Autodesk.Revit.DB.ForgeTypeId)tag.Group;
                var paramType = param.Definition.GetDataType();
#endif

                Guid guid = isShared ? param.GUID : Guid.Empty;

                if (row.Children.Count < 4)
                    continue;

                var valueBox = row.Children[1] as TextBox;  
                var cbValue = row.Children[2] as CheckBox; 
                var cbParam = row.Children[3] as CheckBox;

                if (cbParam?.IsChecked == true)
                {
                    bool? copyVal = cbValue?.IsChecked;    
                    string valStr = valueBox?.Text ?? "";

                    selectedParameters[param] = (isShared, guid, group, copyVal, paramType, isInstance, valStr);
                }
            }

            if (selectedParameters.Count == 0)
            {
                MessageBox.Show($"Не выбрано параметров для копирования");
                return;
            }

            var window = new batchAddingParametersWindowSelectFamily(_uiApp, selectedParameters);
            var revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            new System.Windows.Interop.WindowInteropHelper(window).Owner = revitHandle;
            window.ShowDialog();
        }
    }
}
