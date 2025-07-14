using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using RevitServerAPILib;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.Eventing.Reader;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Xml.Linq;


namespace KPLN_BIMTools_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для batchAddingParametersWindowSelectFamily.xaml
    /// </summary>
    public partial class batchAddingParametersWindowSelectFamily : Window
    {
        private readonly UIApplication _uiApp;
        Dictionary<FamilyParameter, (bool isShared, Guid guid, object Group, bool? CopyValue, object ParamType, bool isInstance, string Value)> _selectedParameters;

        public batchAddingParametersWindowSelectFamily(UIApplication uiApp, 
            Dictionary<FamilyParameter, (bool isShared, Guid guid, object Group, bool? CopyValue, object ParamType, bool isInstance, string Value)> selectedParameters)
        {
            InitializeComponent();

            _uiApp = uiApp;
            _selectedParameters = selectedParameters;
        }

        // Добавление значение в параметр
        public void RelationshipOfValuesWithTypesToAddToParameter(FamilyManager familyManager, FamilyParameter familyParam, String parameterValue, object parameterValueDataType)
        {
            string parameterValueDataTypeValue;

#if (Debug2020 || Revit2020)
            ParameterType paramType20 = (ParameterType)parameterValueDataType;
            parameterValueDataTypeValue = paramType20.ToString();
#endif
#if (Debug2023 || Revit2023)
            ForgeTypeId paramType23 = (ForgeTypeId)parameterValueDataType;
            parameterValueDataTypeValue = paramType23.TypeId;
#endif

            switch (parameterValueDataTypeValue)
            {
                case "autodesk.spec:spec.string-2.0.0":
                case "Text":
                case "autodesk.spec.aec:multilineText-2.0.0":
                case "MultilineText":
                case "autodesk.spec.string:url-2.0.0":
                case "URL":
                    familyManager.Set(familyParam, parameterValue);
                    break;

                case "autodesk.spec:spec.int64-2.0.0":
                case "Integer":
                case "autodesk.spec:spec.bool-1.0.0":
                case "YesNo":
                case "autodesk.spec.aec:numberOfPoles-2.0.0":
                case "NumberOfPoles":
                    if (int.TryParse(parameterValue, out int intBoolValue))
                    {
                        familyManager.Set(familyParam, intBoolValue);
                    }
                    break;

                default:
                if (double.TryParse(parameterValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double dValue))
                {
#if Revit2020 || Debug2020
                    UnitType unitType = familyParam.Definition.UnitType;
                    DisplayUnitType displayUnitType = _uiApp.ActiveUIDocument.Document.GetUnits().GetFormatOptions(unitType).DisplayUnits;
                    double convertedValue = UnitUtils.ConvertToInternalUnits(dValue, displayUnitType);
                    familyManager.Set(familyParam, convertedValue);
#endif
#if Revit2023 || Debug2023
                    ForgeTypeId forgeTypeId = familyParam.Definition.GetDataType();
                    FormatOptions formatOptions = _uiApp.ActiveUIDocument.Document.GetUnits().GetFormatOptions(forgeTypeId);
                    double convertedValue = UnitUtils.ConvertToInternalUnits(dValue, formatOptions.GetUnitTypeId());
                    familyManager.Set(familyParam, convertedValue);
#endif
                }
                    break;
            }           
        }

        // Добавление значений в дебаг-окно
        void AddDebugLine(string text, System.Windows.Media.Color color)
        {
            var paragraph = debugField.Document.Blocks.FirstBlock as Paragraph;
            if (paragraph == null)
            {
                paragraph = new Paragraph();
                debugField.Document.Blocks.Add(paragraph);
            }

            var run = new Run(text + "\n")
            {
                Foreground = new SolidColorBrush(color)
            };
            paragraph.Inlines.Add(run);

            debugField.ScrollToEnd();
        }

        //// XAML. Добавить семейство
        private void BtnAddFamilies_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Revit Family (*.rfa)|*.rfa",
                Multiselect = true,
                Title = "Выберите файлы семейств (*.rfa)"
            };

            if (dialog.ShowDialog() == true)
            {
                foreach (var path in dialog.FileNames)
                {
                    if (!lstFamilyFiles.Items.Contains(path))
                        lstFamilyFiles.Items.Add(path);
                }
            }
        }

        //// XAML. Удалить семейство
        private void BtnRemoveSelected_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in lstFamilyFiles.SelectedItems.Cast<string>().ToList())
                lstFamilyFiles.Items.Remove(item);
        }

        // XAML. Запуск работы плагина
        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            if (lstFamilyFiles.Items.Count == 0)
            {
                MessageBox.Show("Список семейств пуст.\nДобавьте хотя бы один файл семейства для продолжения.", "Внимание", MessageBoxButton.OK,MessageBoxImage.Warning);
                return;
            }

            debugField.Document.Blocks.Clear();

            foreach (string familyPath in lstFamilyFiles.Items)
            {
                Document famDoc = null;
                FamilyManager famMgr = null;
               
                try
                {
                    famDoc = _uiApp.Application.OpenDocumentFile(familyPath);

                    if (chkOpenAfter.IsChecked == true)
                    {
                        _uiApp.OpenAndActivateDocument(familyPath);
                    }

                    if (!famDoc.IsFamilyDocument)
                        continue;

                    AddDebugLine($"ОТКРЫТ ДОКУМЕНТ: {System.IO.Path.GetFileName(familyPath)}", Colors.LimeGreen);

                    famMgr = famDoc.FamilyManager;
                }
                catch (Exception ex)
                {
                    AddDebugLine($"ОШИБКА: {System.IO.Path.GetFileName(familyPath)}: {ex.Message}", Colors.IndianRed);
                }

                using (Transaction t = new Transaction(famDoc, "KPLN. Копирование параметров"))
                {
                    t.Start();

                    bool findError = false;

                    foreach (var kvp in _selectedParameters)
                    {                       
                        FamilyParameter param = kvp.Key;
                        var (isShared, guid, group, copyValue, paramType, isInstance, value) = kvp.Value;

                        string paramName = param.Definition.Name;
                        FamilyParameter existingParam = famMgr.get_Parameter(paramName);

#if (Debug2020 || Revit2020)
                        if (existingParam != null && existingParam.Definition.ParameterType == (ParameterType)paramType && copyValue == false)
#endif
#if (Debug2023 || Revit2023)
                        if (existingParam != null && existingParam.Definition.GetDataType() == (ForgeTypeId)paramType && copyValue == false)
#endif
                        {
                            continue;
                        }

                        FamilyParameter familyParameter = null;
#if (Debug2020 || Revit2020)
                        if (existingParam != null && existingParam.Definition.ParameterType == (ParameterType)paramType)
#endif
#if (Debug2023 || Revit2023)
                        if (existingParam != null && existingParam.Definition.GetDataType() == (ForgeTypeId)paramType)
#endif
                        {
                            familyParameter = existingParam;                           
                        }
                        else
                        {
#if (Debug2020 || Revit2020)
                            if (existingParam != null && existingParam.Definition.ParameterType != (ParameterType)paramType)
                            {
#endif
#if (Debug2023 || Revit2023)
                            if (existingParam != null && existingParam.Definition.GetDataType() != (ForgeTypeId)paramType)
                            {
#endif
                            try
                            {
                                famMgr.RemoveParameter(existingParam);
                            }
                            catch (Exception ex)
                            {
                                AddDebugLine($"ОШИБКА: {paramName}: {ex.Message}", Colors.IndianRed);
                                findError = true;
                            }
                            }

                            if (isShared)
                            {
                                DefinitionFile sharedParameterFile = _uiApp.Application.OpenSharedParameterFile();
                                ExternalDefinition externalDef = null;

                                if (sharedParameterFile != null)
                                {
                                    foreach (DefinitionGroup groupDef in sharedParameterFile.Groups)
                                    {
                                        foreach (Definition def in groupDef.Definitions)
                                        {
                                            if (def is ExternalDefinition extDef && extDef.GUID == guid)
                                            {
                                                externalDef = extDef;
                                                break;
                                            }
                                        }
                                        if (externalDef != null)
                                            break;
                                    }
                                }

                                if (externalDef == null)
                                {
                                    AddDebugLine($"ОШИБКА: {paramName} не найден в текущем ФОП", Colors.IndianRed);
                                    findError = true;
                                }

                                try
                                {
#if (Debug2020 || Revit2020)
                                    familyParameter = famMgr.AddParameter(externalDef, (BuiltInParameterGroup)group, isInstance);
#endif
#if (Debug2023 || Revit2023)
                                    familyParameter = famMgr.AddParameter(externalDef, (ForgeTypeId)group, isInstance);
#endif
                                }
                                catch (Exception ex)
                                {
                                    AddDebugLine($"ОШИБКА: {paramName}: {ex.Message}", Colors.IndianRed);
                                    findError = true;
                                }
                            }
                            else
                            {
                                try
                                {
#if (Debug2020 || Revit2020)
                                    BuiltInParameterGroup group20 = (BuiltInParameterGroup)group;
                                    ParameterType paramType20 = (ParameterType)paramType;
                                    familyParameter = famMgr.AddParameter(paramName, group20, paramType20, isInstance);
#endif
#if (Debug2023 || Revit2023)
                                    ForgeTypeId group23 = (ForgeTypeId)group;
                                    ForgeTypeId paramType23 = (ForgeTypeId)paramType;
                                    familyParameter = famMgr.AddParameter(param.Definition.Name, group23, paramType23, isInstance);
#endif

                                }
                                catch (Exception ex)
                                {
                                    AddDebugLine($"ОШИБКА: {paramName}: {ex.Message}", Colors.IndianRed);
                                    findError = true;
                                }
                            }                          
                        }

                        if (copyValue == true)
                        {
                            try
                            {
                                if (value.StartsWith("="))
                                {
                                    famMgr.SetFormula(familyParameter, value.Substring(1));
                                }
                                else
                                {
                                    RelationshipOfValuesWithTypesToAddToParameter(famMgr, familyParameter, value, paramType);
                                }
                            }
                            catch (Exception ex)
                            {
                                AddDebugLine($"ОШИБКА: {paramName}: {ex.Message}", Colors.IndianRed);
                                findError = true;
                            }
                        }
                    }

                    t.Commit();
                    famDoc.Save();     
                    
                    if (findError)
                    {
                        AddDebugLine($"ДОКУМЕНТ ОБРАБОТАН С ОШИБКАМИ: {System.IO.Path.GetFileName(familyPath)}", Colors.IndianRed);
                    }
                    else
                    {
                        AddDebugLine($"ДОКУМЕНТ ОБРАБОТАН УСПЕШНО: {System.IO.Path.GetFileName(familyPath)}", Colors.LimeGreen);
                    }
                }
               
            }

            AddDebugLine($"ОПЕРАЦИЯ ЗАВЕРШЕНА", Colors.MediumPurple);
        }
    }
}
