//!!!!!!!!!!ВНИМАНИЕ!!!!!!!!!
//Для этого плагина подключен nuget Microsoft.Office.Interop.Excel. Если будешь удалять - удали и этот пакет!!!

#if Revit2020 || Debug2020
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;
using MS.Revit.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using Utils;


namespace MS.Revit.Utils
{
    public class HelperParams
    {
        public static Autodesk.Revit.DB.Parameter GetElemParamInstanseSymbol(Document doc, Element elem, string name)
        {
            return elem.LookupParameter(name) == null ? doc.GetElement(elem.GetTypeId()).LookupParameter(name) : elem.LookupParameter(name);
        }

        public Autodesk.Revit.DB.Parameter GetElemParam(Element elem, string name)
        {
            return this.GetElemParam(elem, name, false);
        }

        public Autodesk.Revit.DB.Parameter GetElemParam(Element elem, string name, bool matchCase)
        {
            StringComparison comparisonType = StringComparison.CurrentCultureIgnoreCase;
            if (matchCase)
                comparisonType = StringComparison.CurrentCulture;
            foreach (Autodesk.Revit.DB.Parameter parameter in elem.Parameters)
            {
                if (parameter.Definition.Name.Equals(name, comparisonType))
                    return parameter;
            }
            return (Autodesk.Revit.DB.Parameter)null;
        }

        public DefinitionFile GetOrCreateSharedParamsFile(Autodesk.Revit.ApplicationServices.Application app)
        {
            string empty = string.Empty;
            try
            {
                string parametersFilename = app.SharedParametersFilename;
                if (string.Empty == parametersFilename)
                {
                    string path = SpecialDirectories.MyDocuments + "\\MyRevitSharedParams.txt";
                    new StreamWriter(path).Close();
                    app.SharedParametersFilename = path;
                }
                return app.OpenSharedParameterFile();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("ERROR: Failed to get or create Shared Params File: " + ex.Message);
                return (DefinitionFile)null;
            }
        }

        public DefinitionGroup GetOrCreateSharedParamsGroup(DefinitionFile defFile, string grpName)
        {
            try
            {
                // Попытка получить группу по имени, если не существует, создать новую
                DefinitionGroup group = defFile.Groups.get_Item(grpName) ?? defFile.Groups.Create(grpName);
                return group;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"ERROR: Failed to get or create Shared Params Group: {ex.Message}", "Error", (MessageBoxButton)MessageBoxButtons.OK, (MessageBoxImage)MessageBoxIcon.Error);
                return null;
            }
        }

        public Definition GetOrCreateSharedParamDefinition(
          DefinitionGroup defGrp,
          string parName,
          bool visible)
        {
            try
            {
                Definition definition = defGrp.Definitions.get_Item(parName);
                if (definition != null)
                {
                    return definition;
                }

                // Создание нового параметра, если его нет
                ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(parName, Autodesk.Revit.DB.ParameterType.Text)
                {
                    Visible = visible
                };

                return defGrp.Definitions.Create(options);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(string.Format("ERROR: Failed to get or create Shared Params Definition: {0}", (object)ex.Message));
                return null;
            }
        }

        public Autodesk.Revit.DB.Parameter GetOrCreateElemSharedParam(
          Element elem,
          string paramName,
          string grpName,
          string paramType,
          bool visible,
          bool instanceBinding)
        {
            try
            {
                Autodesk.Revit.DB.Parameter elemParam = this.GetElemParam(elem, paramName);
                if (elemParam != null)
                    return elemParam;
                HelperParams.BindSharedParamResult sharedParamResult = this.BindSharedParam(elem.Document, elem.Category, paramName, grpName, paramType, visible, instanceBinding);
                return sharedParamResult != HelperParams.BindSharedParamResult.eSuccessfullyBound && sharedParamResult != 0 ? (Autodesk.Revit.DB.Parameter)null : this.GetElemParam(elem, paramName);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(string.Format("Error in getting or creating Element Param: {0}", (object)ex.Message));
                return (Autodesk.Revit.DB.Parameter)null;
            }
        }

        public Autodesk.Revit.DB.Parameter GetOrCreateProjInfoSharedParam(
          Document doc,
          string paramName,
          string grpName,
          string paramType,
          bool visible)
        {
            return this.GetOrCreateElemSharedParam((Element)doc.ProjectInformation, paramName, grpName, paramType, visible, true);
        }

        public HelperParams.BindSharedParamResult BindSharedParam(
          Document doc,
          Category cat,
          string paramName,
          string grpName,
          string paramType,
          bool visible,
          bool instanceBinding)
        {
            try
            {
                Autodesk.Revit.ApplicationServices.Application application = doc.Application;
                CategorySet categorySet = application.Create.NewCategorySet();
                DefinitionBindingMapIterator bindingMapIterator = ((DefinitionBindingMap)doc.ParameterBindings).ForwardIterator();
                while (bindingMapIterator.MoveNext())
                {
                    Definition key = bindingMapIterator.Key;
                    ElementBinding current = (ElementBinding)bindingMapIterator.Current;
                    if (paramName.Equals(key.Name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (current.Categories.Contains(cat))
                        {
                            if (paramType != key.ParameterType.ToString())
                                return HelperParams.BindSharedParamResult.eWrongParamType;
                            if (instanceBinding)
                            {
                                if (current.GetType() != typeof(InstanceBinding))
                                    return HelperParams.BindSharedParamResult.eWrongBindingType;
                            }
                            else if (current.GetType() != typeof(TypeBinding))
                                return HelperParams.BindSharedParamResult.eWrongBindingType;
                            return HelperParams.BindSharedParamResult.eAlreadyBound;
                        }
                        foreach (Category category in current.Categories)
                            categorySet.Insert(category);
                    }
                }
                Definition sharedParamDefinition = this.GetOrCreateSharedParamDefinition(this.GetOrCreateSharedParamsGroup(this.GetOrCreateSharedParamsFile(application), grpName), paramName, visible);
                categorySet.Insert(cat);
                Autodesk.Revit.DB.Binding binding = !instanceBinding ? (Autodesk.Revit.DB.Binding)application.Create.NewTypeBinding(categorySet) : (Autodesk.Revit.DB.Binding)application.Create.NewInstanceBinding(categorySet);
                return ((DefinitionBindingMap)doc.ParameterBindings).Insert(sharedParamDefinition, binding) || doc.ParameterBindings.ReInsert(sharedParamDefinition, binding) ? HelperParams.BindSharedParamResult.eSuccessfullyBound : HelperParams.BindSharedParamResult.eFailed;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(string.Format("Error in Binding Shared Param: {0}", (object)ex.Message));
                return HelperParams.BindSharedParamResult.eFailed;
            }
        }

        public enum BindSharedParamResult
        {
            eAlreadyBound,
            eSuccessfullyBound,
            eWrongParamType,
            eWrongBindingType,
            eFailed,
        }
    }
}


namespace Utils
{
    public class SharedParameter
    {
        public void Add(string paramName, CategorySet catSet, UIApplication uiapp, string groupName)
        {
            UIDocument activeUiDocument = uiapp.ActiveUIDocument;
            Document document = activeUiDocument.Document;

            // Попытка открыть файл общих параметров
            DefinitionFile defFile = activeUiDocument.Application.Application.OpenSharedParameterFile();
            if (defFile == null)
            {
                System.Windows.MessageBox.Show("ERROR: Shared Parameter File is not opened.", "Error", (MessageBoxButton)MessageBoxButtons.OK, (MessageBoxImage)MessageBoxIcon.Error);
                return;
            }

            // Получение группы параметров по имени
            DefinitionGroup defGroup = defFile.Groups.get_Item(groupName);
            if (defGroup == null)
            {
                System.Windows.MessageBox.Show($"ERROR: Group '{groupName}' not found in Shared Parameter File.", "Error", (MessageBoxButton)MessageBoxButtons.OK, (MessageBoxImage)MessageBoxIcon.Error);
                return;
            }

            // Получение параметра по имени
            ExternalDefinition definition = defGroup.Definitions.get_Item(paramName) as ExternalDefinition;
            if (definition == null)
            {
                System.Windows.MessageBox.Show($"ERROR: Parameter '{paramName}' not found in Group '{groupName}'.", "Error", (MessageBoxButton)MessageBoxButtons.OK, (MessageBoxImage)MessageBoxIcon.Error);
                return;
            }

            // Создание привязки экземпляра
            InstanceBinding instanceBinding = uiapp.Application.Create.NewInstanceBinding(catSet);

            // Вставка привязки параметра
            bool result = document.ParameterBindings.Insert(definition, instanceBinding, BuiltInParameterGroup.PG_TEXT);

            if (!result)
            {
                System.Windows.MessageBox.Show($"ERROR: Failed to insert parameter '{paramName}' into the document.", "Error", (MessageBoxButton)MessageBoxButtons.OK, (MessageBoxImage)MessageBoxIcon.Error);
            }
        }
    }

    internal class UnitsProject
    {
        public static bool Check(Document doc)
        {
            Units units = doc.GetUnits();
            if (!(units.DecimalSymbol.ToString() + units.DigitGroupingAmount.ToString() + units.DigitGroupingSymbol.ToString() != "CommaThreeDot"))
                return false;
            TaskDialog.Show("Ошибка!", "Неверно задана группировка десятичных знаков. Зайдите во вкладу управление, единицы проекта и установите значение 123.456.789,00");
            return true;
        }
    }

    public class SharedParameterOrCategories
    {
        public void Add(
          string ParamName,
          string groupName,
          UIApplication uiapp,
          CategorySet newCatSet)
        {
            Document document = uiapp.ActiveUIDocument.Document;
            DefinitionBindingMapIterator bindingMapIterator = ((DefinitionBindingMap)document.ParameterBindings).ForwardIterator();
            int num1 = 0;
            while (bindingMapIterator.MoveNext())
            {
                if (bindingMapIterator.Key.Name.Equals(ParamName))
                    num1 = 1;
            }
            if (num1 == 0)
            {
                SharedParameter sharedParameter = new SharedParameter();
                using (Transaction transaction = new Transaction(document, "Add Binding"))
                {
                    transaction.Start();
                    sharedParameter.Add(ParamName, newCatSet, uiapp, groupName);
                    transaction.Commit();
                }
            }
            else
            {
                HelperParams helperParams = new HelperParams();
                foreach (Category newCat in newCatSet)
                {
                    using (Transaction transaction = new Transaction(document, "Add Binding"))
                    {
                        transaction.Start();
                        int num2 = (int)helperParams.BindSharedParam(document, newCat, ParamName, groupName, ((ParameterType)1).ToString(), true, true);
                        transaction.Commit();
                    }
                }
            }
        }
    }
}

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command_SET_InsulationPipes : IExternalCommand
    {
        public static string commandVersion = "1.5.0.0";
        public Document _doc;
        public List<string> _errorType;
        public IList<Element> _pipeInsulation;
        public IList<Element> _pipeInsulationMeter;
        public IList<Element> _pipeInsulationMeterSquare;
        public bool _checkNestedInsulation = false;
        public List<string> _errorElmentsNotConnectToSystem;
        public Guid _guidParamSizeLanthFitt;
        public Guid _guidParamCount;
        public Guid _guidParamidUnit;
        public Guid _guidParamidType;
        public Guid _guidParamidNamePipeAndDuct;
        public Guid _guidParamidNameing;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication application1 = commandData.Application;
            UIDocument activeUiDocument = application1.ActiveUIDocument;
            this._doc = activeUiDocument.Document;
            this._errorType = new List<string>();
            this._errorElmentsNotConnectToSystem = new List<string>();
            UnitsProject.Check(this._doc);
            this._guidParamSizeLanthFitt = (this._doc.GetElement(new FilteredElementCollector(this._doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((System.Func<Element, bool>)(a => a.Name == "ASML_Размер_Длина фитинга")).FirstOrDefault<Element>().Id) as SharedParameterElement).GuidValue;
            this._guidParamCount = (this._doc.GetElement(new FilteredElementCollector(this._doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((System.Func<Element, bool>)(a => a.Name == "ASML_Количество")).FirstOrDefault<Element>().Id) as SharedParameterElement).GuidValue;
            this._guidParamidNameing = (this._doc.GetElement(new FilteredElementCollector(this._doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((System.Func<Element, bool>)(a => a.Name == "ASML_Наименование")).FirstOrDefault<Element>().Id) as SharedParameterElement).GuidValue;
            this._guidParamidUnit = (this._doc.GetElement(new FilteredElementCollector(this._doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((System.Func<Element, bool>)(a => a.Name == "ASML_Единица измерения")).FirstOrDefault<Element>().Id) as SharedParameterElement).GuidValue;
            this._guidParamidType = (this._doc.GetElement(new FilteredElementCollector(this._doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((System.Func<Element, bool>)(a => a.Name == "ASML_Тип")).FirstOrDefault<Element>().Id) as SharedParameterElement).GuidValue;
            this._guidParamidNamePipeAndDuct = (this._doc.GetElement(new FilteredElementCollector(this._doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((System.Func<Element, bool>)(a => a.Name == "ASML_Наименование трубы и воздуховода")).FirstOrDefault<Element>().Id) as SharedParameterElement).GuidValue;
            this._pipeInsulation = (IList<Element>)((IEnumerable<Element>)new FilteredElementCollector(this._doc).OfClass(typeof(PipeInsulation))).ToList<Element>();
            IList<Element> list1 = this._pipeInsulation
            .Where(a =>
            {
                Element element = this._doc.GetElement(a.GetTypeId());
                string unit = element?.get_Parameter(this._guidParamidUnit)?.AsString();
                return unit != "м2" && unit != "м";
            })
            .ToList();
            if (list1.Count<Element>() > 0)
            {
                TaskDialog.Show("Ошибка!", "Выполнение програмы остановлено т.к. не заполнен параметр ASML_Единица измерения (м или м2) в типах изоляции:\n\n" + string.Join("\n", list1.Select<Element, string>((System.Func<Element, string>)(a => a.Name)).Distinct<string>()));
                return (Result)0;
            }
            using (Transaction transaction = new Transaction(this._doc))
            {
                transaction.Start("Transaction Name");
                if (System.Windows.MessageBox.Show("Перед запуском проверьте и заполните:\nASML_Размер_Длина фитинга\nПРОДОЛЖИТЬ ?", "Предупреждение!", MessageBoxButton.YesNo) == MessageBoxResult.No)
                    return (Result)0;
                if (this._pipeInsulation.Count<Element>() == 0)
                {
                    TaskDialog.Show("Ошибка!", "В проекте отсутствует изоляция труб");
                    return (Result)0;
                }
                ICollection<ElementId> elementIds = (ICollection<ElementId>)new List<ElementId>();
                transaction.Commit();
            }
            this.FindAndCheckInsulation();
            if (this._checkNestedInsulation)
            {
                TaskDialog.Show("Ошибка!", "Проверьте модель плагином Проверка изоляции");
                return (Result)0;
            }
            System.Data.DataTable source1 = new System.Data.DataTable();
            source1.Clear();
            source1.Columns.Add("Type", typeof(string));
            source1.Columns.Add("H", typeof(double));
            source1.Columns.Add("D", typeof(double));
            source1.Columns.Add("Marc", typeof(string));
            source1.Columns.Add("TypeAddName", typeof(string));
            List<List<string>> stringListList = new List<List<string>>();
            string Filename1 = @"Y:\Жилые здания\Самолет Сетунь\6.BIM\11.Автоматизация_Самолет\ОВВК\Спецификация типов изоляции труб из каталога.xlsx";
            Microsoft.Office.Interop.Excel.Application application2 = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook;
            workbook = application2.Workbooks.Open(Filename1, (object)0, (object)true, (object)5, (object)"", (object)"", (object)true, (object)XlPlatform.xlWindows, (object)"\t", (object)false, (object)false, (object)0, (object)true, (object)1, (object)0);
            Microsoft.Office.Interop.Excel.Range usedRange = ((_Worksheet)workbook.Sheets[(object)1]).UsedRange;
            int count = usedRange.Rows.Count;
            object[,] objArray = (object[,])usedRange.get_Value((object)XlRangeValueDataType.xlRangeValueDefault);
            for (int index = 2; index < count + 1; ++index)
            {
                if (objArray[index, 1] != null && objArray[index, 1].ToString().Length > 2)
                {
                    DataRow row = source1.NewRow();
                    try
                    {
                        row["Type"] = (object)objArray[index, 1].ToString();
                        row["H"] = (object)double.Parse(objArray[index, 2].ToString());
                        row["D"] = (object)double.Parse(objArray[index, 3].ToString());
                        row["Marc"] = (object)objArray[index, 4].ToString();
                        object itemObjArr = (object)objArray[index, 5];
                        if (itemObjArr != null)
                            row["TypeAddName"] = itemObjArr.ToString();
                        else
                            row["TypeAddName"] = (object)"";
                    }
                    catch
                    {
                        TaskDialog.Show("Ошибка!", "Проверьте строку " + index.ToString() + " в файле Спецификация типов изоляции труб из каталога.xlsx на правильность заполнения данными, заполнение параметров приостановлено. Для исправления ошибки обратитесь в BIM отдел.");
                        return (Result)0;
                    }
                    source1.Rows.Add(row);
                }
            }
            workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            application2.Quit();
            try
            {
                Element element = this._pipeInsulation.FirstOrDefault();
                if (element != null)
                {
                    Autodesk.Revit.DB.Parameter param = element.get_Parameter(this._guidParamCount);
                    if (param != null && !param.IsReadOnly)
                        param.Set(1);
                }
            }
            catch
            {
                List<string> stringList1 = new List<string>()
                {
                    "ASML_Количество"
                };
                List<string> stringList2 = new List<string>()
                {
                  "Материалы изоляции труб"
                };
                SharedParameterOrCategories parameterOrCategories = new SharedParameterOrCategories();
                Categories categories = this._doc.Settings.Categories;
                CategorySet newCatSet = activeUiDocument.Application.Application.Create.NewCategorySet();
                foreach (string str in stringList2)
                    if (categories is CategoryNameMap categoryMap)
                    {
                        Category category = categoryMap.get_Item(str);
                        if (category != null)
                            newCatSet.Insert(category);
                        else
                            System.Windows.MessageBox.Show($"Category '{str}' not found in CategoryNameMap.", "Error", (MessageBoxButton)MessageBoxButtons.OK, (MessageBoxImage)MessageBoxIcon.Error);
                    }
                if (activeUiDocument.Application.Application.OpenSharedParameterFile() == null)
                {
                    TaskDialog.Show("Ошибка!", "Загрузите ФОП (файл общих параметров) в проект");
                    return (Result)0;
                }
                foreach (string ParamName in stringList1)
                    parameterOrCategories.Add(ParamName, "28_ИОС_Общие", application1, newCatSet);
            }
            ICollection<ElementId> source2 = (ICollection<ElementId>)new List<ElementId>();
            using (Transaction transaction = new Transaction(this._doc))
            {
                transaction.Start("Transaction Name");

                foreach (Element element in this._pipeInsulationMeter)
                {
                    if (element != null && element.get_Parameter(this._guidParamSizeLanthFitt)?.AsDouble() == 0.0)
                    {
                        var insulation = element as InsulationLiningBase;
                        if (insulation != null)
                        {
                            source2.Add(insulation.HostElementId);
                        }
                    }
                    else
                    {
                        try
                        {
                            double sizeLanthFitt = element.get_Parameter(this._guidParamSizeLanthFitt)?.AsDouble() ?? 0.0;
                            double newValue = sizeLanthFitt / (1250.0 / 381.0);

                            element.get_Parameter(this._guidParamCount)?.Set(newValue);
                        }
                        catch (Exception ex)
                        {
                            var insulation = element as InsulationLiningBase;
                            if (insulation != null)
                            {
                                source2.Add(insulation.HostElementId);
                            }
                            System.Windows.MessageBox.Show($"Error updating element {element.Id}: {ex.Message}", "Error", (MessageBoxButton)MessageBoxButtons.OK, (MessageBoxImage)MessageBoxIcon.Error);
                        }
                    }
                }

                foreach (Element element in this._pipeInsulationMeterSquare)
                {
                    if (element != null && element.get_Parameter(this._guidParamSizeLanthFitt)?.AsDouble() == 0.0)
                    {
                        var insulation = element as InsulationLiningBase;
                        if (insulation != null)
                        {
                            source2.Add(insulation.HostElementId);
                        }
                    }
                    else
                    {
                        try
                        {
                            double sizeLanthFitt = element.get_Parameter(this._guidParamSizeLanthFitt)?.AsDouble() ?? 0.0;

                            var insulation = element as InsulationLiningBase;
                            if (insulation != null)
                            {
                                double hostSize = this._doc.GetElement(insulation.HostElementId)
                                    .get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)?.AsDouble() ?? 0.0;
                                double insulationParam = element.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_PIPE)?.AsDouble() ?? 0.0;

                                double newValue = sizeLanthFitt / (1250.0 / 381.0) * (hostSize / (1250.0 / 381.0) + insulationParam / (1250.0 / 381.0) * 2.0) * Math.PI;

                                element.get_Parameter(this._guidParamCount)?.Set(newValue);
                            }
                        }
                        catch (Exception ex)
                        {
                            var insulation = element as InsulationLiningBase;
                            if (insulation != null)
                            {
                                source2.Add(insulation.HostElementId);
                            }
                            System.Windows.MessageBox.Show($"Error updating element {element.Id}: {ex.Message}", "Error", (MessageBoxButton)MessageBoxButtons.OK, (MessageBoxImage)MessageBoxIcon.Error);
                        }
                    }
                }

                transaction.Commit();
            }
            if (source2.Count<ElementId>() > 0 && System.Windows.MessageBox.Show("В " + source2.Count<ElementId>().ToString() + " элементах изоляции не заполнен параметр ASML_Размер_Длина фитинга или ASML_Размер_Длина фитинга равен нулю\nвыделить элементы на которых располагается изоляция в модели и остановить выполнение программы?", "Предупреждение!", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                new UIDocument(this._doc).Selection.SetElementIds(source2);
                return (Result)0;
            }
            using (Transaction transaction = new Transaction(this._doc))
            {
                transaction.Start("Transaction Name");
                ICollection<ElementId> source3 = (ICollection<ElementId>)new List<ElementId>();
                List<string> source4 = new List<string>();
                List<Element> elementList = new List<Element>();
                List<Element> list2;
                try
                {
                    list2 = this._pipeInsulation.Where<Element>((System.Func<Element, bool>)(a => this._doc.GetElement((a as InsulationLiningBase).HostElementId).Category.Name != "Трубы")).Where<Element>((System.Func<Element, bool>)(a => this._doc.GetElement((a as InsulationLiningBase).HostElementId).Category.Name != "Соединительные детали трубопроводов")).Select<Element, Element>((System.Func<Element, Element>)(a => a)).ToList<Element>();
                }
                catch
                {
                    TaskDialog.Show("Предупреждение!", "В модели присутствуют экземпляры изоляции трубопровода без основы. Проверьте и удалите изоляцию без основы плагином по проверке изоляции.Заполнение параметров приостановлено.");
                    return (Result)0;
                }
                if (list2.Count<Element>() > 0)
                {
                    foreach (Element element in list2)
                    {
                        source3.Add(element.Id);
                        source4.Add(this._doc.GetElement((element as InsulationLiningBase).HostElementId).Name);
                    }
                    List<string> list3 = source4.Distinct<string>().ToList<string>();
                    TaskDialog.Show("Предупреждение!", "Проверьте что изоляция труб назначена только категориям труб и соединительным деталям труб, возможны исключения (пример: Петлевые компенсаторы)\nКоличество=" + source3.Count<ElementId>().ToString() + "\n\nИмена семейств:\n" + string.Join("\n", (IEnumerable<string>)list3));
                }
                this._pipeInsulation.Select<Element, string>((System.Func<Element, string>)(a => this._doc.GetElement((a as InsulationLiningBase).HostElementId).Category.Name)).Distinct<string>().ToList<string>();
                List<string> list4 = this._pipeInsulationMeter.Select<Element, string>((System.Func<Element, string>)(a => a.Name)).Distinct<string>().ToList<string>();
                List<string> source5 = new List<string>();
                foreach (string str in list4)
                {
                    string TypeInsuli = str;
                    this._pipeInsulationMeter.Where<Element>((System.Func<Element, bool>)(a => a.Name == TypeInsuli)).ToList<Element>();
                    System.Data.DataTable dataTable = new System.Data.DataTable();
                    try
                    {
                        dataTable = source1.AsEnumerable().Where<DataRow>((System.Func<DataRow, bool>)(r => r.Field<string>("Type") == TypeInsuli)).CopyToDataTable<DataRow>();
                    }
                    catch
                    {
                        source5.Add(TypeInsuli);
                    }
                }
                if (source5.Count<string>() > 0)
                {
                    TaskDialog.Show("Ошибка!", "Выполнение программы остановлено т.к. в таблице типов отсутствуют следующие типы изоляции:\n\n" + string.Join("\n", source5.Distinct<string>()) + "\n\nОбратитесь в BIM отдел для добавления типоразмеров и повторите запуск плагина");
                    return (Result)0;
                }
                foreach (string str in list4)
                {
                    string TypeInsuli = str;
                    List<Element> list5 = this._pipeInsulationMeter.Where<Element>((System.Func<Element, bool>)(a => a.Name == TypeInsuli)).ToList<Element>();
                    System.Data.DataTable dataTable1 = new System.Data.DataTable();
                    System.Data.DataTable dataTable2;
                    try
                    {
                        dataTable2 = source1.AsEnumerable().Where<DataRow>((System.Func<DataRow, bool>)(r => r.Field<string>("Type") == TypeInsuli)).CopyToDataTable<DataRow>();
                    }
                    catch
                    {
                        continue;
                    }
                    foreach (Element Elem in list5)
                        this.SetInsulation(this._doc, Elem, dataTable2, 0.0, 0.0);
                }
                transaction.Commit();
            }
            
            if (this._errorType.Count<string>() > 0 && System.Windows.MessageBox.Show("Не найден подходящий размер изоляции, обратитесь в BIM отдел для добавления типоразмеров и повторите запуск плагина\n\nСписок труб, для которых не найден тип изоляции:\n\n" + string.Join("\n", this._errorType.Distinct<string>()) + "\n\nПродолжить выполнение плагина?", "Ошибка!", MessageBoxButton.YesNo) == MessageBoxResult.No)
                return (Result)0;
            
            using (Transaction transaction = new Transaction(this._doc))
            {
                transaction.Start("Transaction Name");

                // Получение уникальных имен из коллекции
                var distinctNames = this._pipeInsulationMeterSquare
                    .Select(a => a.Name)
                    .Distinct()
                    .ToList();

                foreach (string typeInsuli in distinctNames)
                {
                    // Получаем элементы с тем же именем
                    var elementsWithSameName = this._pipeInsulationMeterSquare
                        .Where(a => a.Name == typeInsuli)
                        .ToList();

                    foreach (Element element in elementsWithSameName)
                    {
                        // Получение необходимых параметров
                        var hostElement = this._doc.GetElement((element as InsulationLiningBase)?.HostElementId);
                        var paramThickness = element.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_PIPE)?.AsDouble() * 304.8;
                        var paramDiameter = hostElement?.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.AsDouble() * 304.8;
                        var paramName = this._doc.GetElement(element.GetTypeId())
                            ?.get_Parameter(this._guidParamidNamePipeAndDuct)?.AsString();

                        if (paramThickness != null && paramDiameter != null && paramName != null)
                        {
                            string newName = $"{paramName} толщиной {paramThickness:F2} мм для труб Ду {Math.Round((double)paramDiameter, 2)} мм";
                            element.get_Parameter(this._guidParamidNameing)?.Set(newName);
                        }
                    }
                }

                // Разделение на элементы с и без "ТолщинаИзоляции"
                var list6 = this._pipeInsulationMeterSquare
                    .Where(a => !this._doc.GetElement(a.GetTypeId())
                    .get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsString().Contains("ТолщинаИзоляции") ?? false)
                    .ToList();

                var list7 = this._pipeInsulationMeterSquare
                    .Where(a => this._doc.GetElement(a.GetTypeId())
                    .get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsString().Contains("ТолщинаИзоляции") ?? false)
                    .ToList();

                // Обновление для элементов без "ТолщинаИзоляции"
                foreach (Element element in list6)
                {
                    var paramType = this._doc.GetElement(element.GetTypeId())
                        ?.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsString();
                    element.get_Parameter(this._guidParamidType)?.Set(paramType);
                }

                // Обновление для элементов с "ТолщинаИзоляции"
                foreach (Element element in list7)
                {
                    var paramType = this._doc.GetElement(element.GetTypeId())
                        ?.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)?.AsString();

                    if (!string.IsNullOrEmpty(paramType))
                    {
                        string beforeThickness = paramType.Substring(0, paramType.IndexOf("ТолщинаИзоляции"));
                        string afterThickness = paramType.Substring(paramType.IndexOf("ТолщинаИзоляции") + 15);
                        var paramThickness = element.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_PIPE)?.AsDouble() * 304.8;

                        if (paramThickness != null)
                        {
                            string newType = beforeThickness + paramThickness.ToString() + afterThickness;
                            element.get_Parameter(this._guidParamidType)?.Set(newType);
                        }
                    }
                }

                transaction.Commit();
            }
            List<ElementId> elementIdList = new List<ElementId>();
            using (Transaction transaction = new Transaction(this._doc))
            {
                transaction.Start("Transaction Name");
                List<Element> list8 = this._pipeInsulation.Where<Element>((System.Func<Element, bool>)(a => this._doc.GetElement((a as InsulationLiningBase).HostElementId).Category.Name == "Соединительные детали трубопроводов")).Select<Element, Element>((System.Func<Element, Element>)(a => a)).ToList<Element>().Concat<Element>((IEnumerable<Element>)this._pipeInsulation.Where<Element>((System.Func<Element, bool>)(a => this._doc.GetElement((a as InsulationLiningBase).HostElementId).Category.Name == "Арматура трубопроводов")).Select<Element, Element>((System.Func<Element, Element>)(a => a)).ToList<Element>()).ToList<Element>();
                ElementClassFilter elementClassFilter = new ElementClassFilter(typeof(PipeInsulation));
                ICollection<string> source6 = (ICollection<string>)new List<string>();
                foreach (Element element1 in list8)
                {
                    Element elem = element1;
                    FamilyInstance element2 = this._doc.GetElement((elem as InsulationLiningBase).HostElementId) as FamilyInstance;
                    List<Element> source7 = new List<Element>();
                    List<Connector> source8 = new List<Connector>();
                    foreach (Connector connector in element2.MEPModel.ConnectorManager.Connectors)
                    {
                        foreach (Connector allRef in connector.AllRefs)
                        {
                            if (allRef.Owner.Id.ToString() != ((Element)element2).Id.ToString())
                            {
                                source7.Add(allRef.Owner);
                                source8.Add(allRef);
                            }
                        }
                    }
                    double diamOut = 0.0;
                    double num1;
                    try
                    {
                        num1 = source8.Where<Connector>((System.Func<Connector, bool>)(a => a.Domain.ToString() == "DomainPiping")).Select<Connector, double>((System.Func<Connector, double>)(a => a.Radius)).Max() * 2.0;
                    }
                    catch
                    {
                        this._errorElmentsNotConnectToSystem.Add(((Element)element2).Id.ToString());
                        continue;
                    }
                    List<Element> list9 = source7.Where<Element>((System.Func<Element, bool>)(a => a.Category.Name == "Соединительные детали трубопроводов" || a.Category.Name == "Арматура трубопроводов" || a.Category.Name == "Трубы")).ToList<Element>();
                    if (list9.Count<Element>() > 0)
                    {
                        try
                        {
                            diamOut = Command_SET_InsulationPipes.FindSizeNextPipe(list9, num1);
                        }
                        catch
                        {
                            this._errorElmentsNotConnectToSystem.Add(((Element)element2).Id.ToString());
                            continue;
                        }
                    }
                    list9.Clear();
                    System.Data.DataTable newDt = new System.Data.DataTable();
                    try
                    {
                        newDt = source1.AsEnumerable().Where<DataRow>((System.Func<DataRow, bool>)(r => r.Field<string>("Type") == elem.Name)).CopyToDataTable<DataRow>();
                    }
                    catch
                    {
                    }
                    if (this._doc.GetElement(elem.GetTypeId()).get_Parameter(this._guidParamidUnit).AsString() == "м2")
                    {
                        if (!this._doc.GetElement(elem.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString().Contains("ТолщинаИзоляции"))
                            elem.get_Parameter(this._guidParamidType).Set(this._doc.GetElement(elem.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString());
                        double num2;
                        if (this._doc.GetElement(elem.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString().Contains("ТолщинаИзоляции"))
                        {
                            string str4 = this._doc.GetElement(elem.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString();
                            string str5 = str4.Substring(0, str4.IndexOf("ТолщинаИзоляции"));
                            string str6 = str4.Substring(str4.IndexOf("ТолщинаИзоляции") + 15);
                            Autodesk.Revit.DB.Parameter parameter = elem.get_Parameter(this._guidParamidType);
                            string str7 = str5;
                            num2 = elem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsDouble() * 304.8;
                            string str8 = num2.ToString();
                            string str9 = str6;
                            string str10 = str7 + str8 + str9;
                            parameter.Set(str10);
                        }
                        Autodesk.Revit.DB.Parameter parameter1 = elem.get_Parameter(this._guidParamidNameing);
                        string[] strArray = new string[6]
                        {
                          this._doc.GetElement(elem.GetTypeId()).get_Parameter(this._guidParamidNamePipeAndDuct).AsString(),
                          " толщиной ",
                          null,
                          null,
                          null,
                          null
                        };
                        num2 = elem.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_PIPE).AsDouble() * 304.8;
                        strArray[2] = num2.ToString();
                        strArray[3] = " мм для труб Ду";
                        num2 = Math.Round(num1 * 304.8, 2);
                        strArray[4] = num2.ToString();
                        strArray[5] = " мм";
                        string str = string.Concat(strArray);
                        parameter1.Set(str);
                        elem.get_Parameter(this._guidParamCount).Set(elem.get_Parameter(this._guidParamSizeLanthFitt).AsDouble() / (1250.0 / 381.0) * (diamOut / (1250.0 / 381.0) + elem.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_PIPE).AsDouble() / (1250.0 / 381.0) * 2.0) * Math.PI);
                    }
                    if (this._doc.GetElement(elem.GetTypeId()).get_Parameter(this._guidParamidUnit).AsString() == "м")
                    {
                        this.SetInsulation(this._doc, elem, newDt, num1, diamOut);
                        elem.get_Parameter(this._guidParamCount).Set(elem.get_Parameter(this._guidParamSizeLanthFitt).AsDouble() / (1250.0 / 381.0));
                    }
                }
                if (source6.Count<string>() > 0)
                    TaskDialog.Show("Ошибка!", "Обратитесь в отдел BIM");
                transaction.Commit();
            }
            using (Transaction transaction = new Transaction(this._doc))
            {
                transaction.Start("Transaction Name");
                new UIDocument(this._doc).Selection.SetElementIds((ICollection<ElementId>)elementIdList);
                transaction.Commit();
            }
            if (this._errorElmentsNotConnectToSystem.Count > 0)
            {
                TaskDialog.Show("Ошибка!", "Невозможно заполнить параметры для изоляции элементов, которые не подключены к системам с трубами. Подключите элементы к трубе или перезапустите плагин. При невозможности соединения или повторной ошибке - заполните параметры вручную. Id элементов см. в файле File.txt");
                File.WriteAllBytes(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\File.txt", new byte[0]);
                File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\File.txt", string.Join("\n", (IEnumerable<string>)this._errorElmentsNotConnectToSystem));
                new Process()
                {
                    StartInfo = 
                    {
                        FileName = "notepad.exe",
                        Arguments = (Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\File.txt")
                    }
                }.Start();
                return (Result)0;
            }
            if (this._errorType.Count<string>() > 0 && System.Windows.MessageBox.Show("Не найден подходящий размер изоляции, обратитесь в BIM отдел для добавления типоразмеров и повторите запуск плагина\n\nСписок труб, для которых не найден тип изоляции:\n\n" + string.Join("\n", this._errorType.Distinct<string>()) + "\n\nПродолжить выполнение плагина?", "Ошибка!", MessageBoxButton.YesNo) == MessageBoxResult.No)
                return (Result)0;
            this._errorType.Clear();
            TaskDialog.Show("S.ОВВК", "Заполнение параметров изоляции труб выполнено!");
            return Result.Succeeded;
        }

        public void FindAndCheckInsulation()
        {
            try
            {
                this._pipeInsulationMeter = (IList<Element>)this._pipeInsulation.Where<Element>((System.Func<Element, bool>)(a => this._doc.GetElement(a.GetTypeId()).get_Parameter(this._guidParamidUnit).AsString() == "м")).Where<Element>((System.Func<Element, bool>)(a => this._doc.GetElement((a as InsulationLiningBase).HostElementId).Category.Name == "Трубы")).ToList<Element>();
                this._pipeInsulationMeterSquare = (IList<Element>)this._pipeInsulation.Where<Element>((System.Func<Element, bool>)(a => this._doc.GetElement(a.GetTypeId()).get_Parameter(this._guidParamidUnit).AsString() == "м2")).Where<Element>((System.Func<Element, bool>)(a => this._doc.GetElement((a as InsulationLiningBase).HostElementId).Category.Name == "Трубы")).ToList<Element>();
            }
            catch
            {
                this._checkNestedInsulation = true;
            }
        }

        private static double FindSizeNextPipe(List<Element> elementsFitt, double maxDiamConnector)
        {
            int num = 0;
            List<Element> source = new List<Element>();
            if (elementsFitt.Where<Element>((System.Func<Element, bool>)(a => a.Category.Name == "Трубы")).ToList<Element>().Count<Element>() > 0)
            {
                source = elementsFitt;
            }
            else
            {
                do
                {
                    ++num;
                    foreach (Element element in elementsFitt)
                    {
                        ConnectorSet connectorSet = (ConnectorSet)null;
                        if (element.Category.Name == "Трубы")
                            connectorSet = (element as MEPCurve).ConnectorManager.Connectors;
                        else if (element.Category.Name == "Соединительные детали трубопроводов" || element.Category.Name == "Арматура трубопроводов")
                            connectorSet = (element as FamilyInstance).MEPModel.ConnectorManager.Connectors;
                        if (connectorSet != null)
                        {
                            foreach (Connector connector in connectorSet)
                            {
                                foreach (Connector allRef in connector.AllRefs)
                                {
                                    try
                                    {
                                        if (allRef.Owner.Category.Name == "Трубы")
                                            source.Add(allRef.Owner);
                                    }
                                    catch
                                    {
                                    }
                                }
                            }
                        }
                    }
                    elementsFitt = source.ToList<Element>();
                }
                while (num != 20 && source.Where<Element>((System.Func<Element, bool>)(a => a.Category.Name == "Трубы")).ToList<Element>().Count<Element>() == 0);
            }
            List<MEPSize> list = ((Segment)(source.Where<Element>((System.Func<Element, bool>)(a => a.Category.Name == "Трубы")).ToList<Element>().First<Element>() as Pipe).PipeSegment).GetSizes().ToList<MEPSize>();
            double sizeNextPipe = maxDiamConnector;
            if (list.Where<MEPSize>((System.Func<MEPSize, bool>)(a => a.NominalDiameter == maxDiamConnector)).Count<MEPSize>() > 0)
                sizeNextPipe = list.Where<MEPSize>((System.Func<MEPSize, bool>)(a => a.NominalDiameter == maxDiamConnector)).First<MEPSize>().OuterDiameter;
            return sizeNextPipe;
        }

        private static void Msg(string s) => TaskDialog.Show("Revit", s);

        public void SetInsulation(
          Document doc,
          Element Elem,
          System.Data.DataTable newDt,
          double diamNominal,
          double diamOut)
        {
            if (diamOut == 0.0)
            {
                Autodesk.Revit.DB.Parameter paramData = doc.GetElement((Elem as InsulationLiningBase).HostElementId).get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                if (paramData == null)
                    return;
                else
                    diamOut = paramData.AsDouble();
            }
            string str1 = "Ду";
            string str2 = "Дн";
            string H = (Elem.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_PIPE).AsDouble() * 304.8).ToString();
            DataView defaultView = newDt.DefaultView;
            defaultView.Sort = "D";
            System.Data.DataTable table = defaultView.ToTable();
            double newListT2 = 0.0;
            try
            {
                newListT2 = table.AsEnumerable().Where<DataRow>((System.Func<DataRow, bool>)(r => r.Field<double>("D") >= diamOut * 304.8)).Where<DataRow>((System.Func<DataRow, bool>)(r => r.Field<double>("H").ToString() == H)).Select<DataRow, double>((System.Func<DataRow, double>)(a => a.Field<double>("D"))).ToList<double>().Min();
            }
            catch
            {
                this._errorType.Add("Наружный диаметр-" + (diamOut * 304.8).ToString() + " толщина изоляции " + H.ToString() + "мм, тип изоляции = " + Elem.Name);
                return;
            }
            List<string> list = table.AsEnumerable().Where<DataRow>((System.Func<DataRow, bool>)(r => r.Field<double>("D") == newListT2)).Where<DataRow>((System.Func<DataRow, bool>)(r => r.Field<double>("H").ToString() == H)).Select<DataRow, string>((System.Func<DataRow, string>)(a => a.Field<string>("Marc"))).ToList<string>();
            string str3 = table.AsEnumerable().Where<DataRow>((System.Func<DataRow, bool>)(r => r.Field<double>("D") >= diamOut * 304.8)).Where<DataRow>((System.Func<DataRow, bool>)(r => r.Field<double>("H").ToString() == H)).Select<DataRow, string>((System.Func<DataRow, string>)(a => a.Field<string>("TypeAddName"))).ToList<string>().First<string>();
            if (str3.Length > 2)
                str3 = " " + str3;
            if (list.Count<string>() <= 0)
                return;
            Elem.get_Parameter(this._guidParamidType).Set(list.First<string>());
            try
            {
                if (diamNominal == 0.0)
                    diamNominal = doc.GetElement((Elem as InsulationLiningBase).HostElementId).get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble();
                double num;
                if (doc.GetElement(Elem.GetTypeId()).get_Parameter(BuiltInParameter.KEYNOTE_PARAM).AsString() != "Без толщины стенки")
                {
                    Elem.get_Parameter(this._guidParamidNameing).Set(doc.GetElement(Elem.GetTypeId()).get_Parameter(this._guidParamidNamePipeAndDuct).AsString() + str3 + " толщиной " + H + " мм для труб " + str1 + Math.Round(diamNominal * 304.8, 2).ToString() + " мм");
                }
                else
                {
                    Autodesk.Revit.DB.Parameter parameter = Elem.get_Parameter(this._guidParamidNameing);
                    string[] strArray = new string[5]
                    {
                        doc.GetElement(Elem.GetTypeId()).get_Parameter(this._guidParamidNamePipeAndDuct).AsString(),
                        " для труб ",
                        str1,
                        null,
                        null
                    };
                    num = Math.Round(diamNominal * 304.8, 2);
                    strArray[3] = num.ToString();
                    strArray[4] = " мм";
                    string str4 = string.Concat(strArray);
                    parameter.Set(str4);
                }
                if (doc.GetElement(Elem.GetTypeId()).get_Parameter(BuiltInParameter.KEYNOTE_PARAM).AsString() == "Дн")
                {
                    Autodesk.Revit.DB.Parameter parameter = Elem.get_Parameter(this._guidParamidNameing);
                    string[] strArray = new string[8]
                    {
                        doc.GetElement(Elem.GetTypeId()).get_Parameter(this._guidParamidNamePipeAndDuct).AsString(),
                        str3,
                        " толщиной ",
                        H,
                        " мм для труб ",
                        str2,
                        null,
                        null
                    };
                    num = Math.Round(diamNominal * 304.8, 2);
                    strArray[6] = num.ToString();
                    strArray[7] = " мм";
                    string str5 = string.Concat(strArray);
                    parameter.Set(str5);
                }
            }
            catch
            {
            }
        }
    }
}
#endif
