using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class CommandTitleBlockChanger : IExternalCommand
    {
        private static Dictionary<ElementId, int[]> _sheetParamsDict;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            UserSelectAlgoritm userSelectForm = new UserSelectAlgoritm();
            userSelectForm.ShowDialog();
            switch (userSelectForm.UserAlgoritm)
            {
                case Algoritm.Close:
                    return Result.Cancelled;
                case Algoritm.SaveSize:
                    SaveSize(doc);
                    return Result.Succeeded;
                case Algoritm.RevalueSize:
                    RevalueSize(doc);
                    return Result.Succeeded;
                case Algoritm.LoadParams:
                    LoadParams(doc);
                    return Result.Succeeded;
            }

            return Result.Succeeded;
        }

        private void SaveSize(Document doc)
        {
            Element[] docSheetsArr = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Sheets)
                .WhereElementIsNotElementType()
                .ToArray();
            _sheetParamsDict = new Dictionary<ElementId, int[]>(docSheetsArr.Length);
            
            for (int i = 0; i < docSheetsArr.Length; i++)
            {
                Element elemView = docSheetsArr[i];
                if (elemView is View view)
                {
                    Element[] sheetTBlockssArr = new FilteredElementCollector(doc, view.Id)
                        .OfCategory(BuiltInCategory.OST_TitleBlocks)
                        .WhereElementIsNotElementType()
                        .ToArray();
                    foreach (FamilyInstance fiTBlock in sheetTBlockssArr)
                    {
                        if (fiTBlock.Symbol.FamilyName.ToLower().Contains("основнаянадпись"))
                        {
                            Parameter formatParam = fiTBlock.LookupParameter("Формат А");
                            Parameter multiplicityParam = fiTBlock.LookupParameter("Кратность");
                            if (formatParam == null || multiplicityParam == null)
                                throw new Exception($"Элемент {fiTBlock.Id} - не имеет нужных параметров! Обратись к разработчику");
                            else
                            {
                                int formatParamData = formatParam.AsInteger();
                                int multiplicityParamData = multiplicityParam.AsInteger();
                                _sheetParamsDict[view.Id] = new int[2] { formatParamData, multiplicityParamData };
                            }
                        }
                    }
                }
                else
                    throw new Exception($"Элемент {elemView.Id} - не лист");
            }
        }

        private void RevalueSize(Document doc)
        {
            if (_sheetParamsDict == null)
            {
                TaskDialog taskDialog = new TaskDialog("ОШИБКА")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                    MainContent = "Сначала нужно запомнить параметры листов, а потом уже их менять!",
                };
                taskDialog.Show();
                return;
            }

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("KPLN: Замена");
                
                foreach (KeyValuePair<ElementId, int[]> kvp in _sheetParamsDict)
                {
                    Element[] sheetTBlockssArr = new FilteredElementCollector(doc, kvp.Key)
                        .OfCategory(BuiltInCategory.OST_TitleBlocks)
                        .WhereElementIsNotElementType()
                        .ToArray();
                    foreach (FamilyInstance fiTBlock in sheetTBlockssArr)
                    {
                        if (fiTBlock.Symbol.FamilyName.ToLower().Contains("основные надписи"))
                        {
                            Parameter formatParam = fiTBlock.LookupParameter("А");
                            Parameter multiplicityParam = fiTBlock.LookupParameter("х");
                            if (formatParam == null || multiplicityParam == null)
                            {
                                trans.RollBack();
                                throw new Exception($"Элемент {fiTBlock.Id} - не имеет нужных параметров! Обратись к разработчику");
                            }
                            else
                            {
                                formatParam.Set(kvp.Value[0]);
                                multiplicityParam.Set(kvp.Value[1]);
                            }
                        }
                    }
                }

                trans.Commit();
            }
        }

        private void LoadParams(Document doc)
        {
            doc.Application.SharedParametersFilename = @"X:\BIM\4_ФОП\КП_Файл общих парамеров.txt";
            
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("KPLN: Параметры");

                // Отметка
                Definition sht_absElev = GetParameter_ByGroupAndName(doc, "09 Заполнение штампа", "SHT_Абсолютная отметка");
                SetBinding(doc, Category.GetCategory(doc, BuiltInCategory.OST_ProjectInformation), sht_absElev, true);

                // Инфо о проекте
                Definition sht_buildType = GetParameter_ByGroupAndName(doc, "09 Заполнение штампа", "SHT_Вид строительства");
                SetBinding(doc, Category.GetCategory(doc, BuiltInCategory.OST_ProjectInformation), sht_buildType, true);

                // Номер листа
                Definition kpSheetNumb = GetParameter_ByGroupAndName(doc, "09 Заполнение штампа", "КП_Ш_Номер листа");
                SetBinding(doc, Category.GetCategory(doc, BuiltInCategory.OST_Sheets), kpSheetNumb, true);

                // Номер листа сквозной
                Definition kpSheetNumb_All = GetParameter_ByGroupAndName(doc, "09 Заполнение штампа", "КП_Ш_Номер листа сквозной");
                SetBinding(doc, Category.GetCategory(doc, BuiltInCategory.OST_Sheets), kpSheetNumb_All, true);

                // Группа листов
                Definition kpSheetGroup = GetParameter_ByGroupAndName(doc, "05 Необязательные ОБЩИЕ", "Орг.ГруппаЛистов");
                SetBinding(doc, Category.GetCategory(doc, BuiltInCategory.OST_Sheets), kpSheetGroup, true);

                // Замечания к листу
                Definition kpSheetError = GetParameter_ByGroupAndName(doc, "05 Необязательные ОБЩИЕ", "Орг.ЗамечаниеКЛисту");
                SetBinding(doc, Category.GetCategory(doc, BuiltInCategory.OST_Sheets), kpSheetError, true);

                // Колво Уч1
                Definition kpShCU1 = GetParameter_ByGroupAndName(doc, "09 Заполнение штампа", "Ш.КолвоУч1Текст");
                SetBinding(doc, Category.GetCategory(doc, BuiltInCategory.OST_Sheets), kpShCU1, true);

                // Колво Уч2
                Definition kpShCU2 = GetParameter_ByGroupAndName(doc, "09 Заполнение штампа", "Ш.КолвоУч2Текст");
                SetBinding(doc, Category.GetCategory(doc, BuiltInCategory.OST_Sheets), kpShCU2, true);

                // Колво Уч3
                Definition kpShCU3 = GetParameter_ByGroupAndName(doc, "09 Заполнение штампа", "Ш.КолвоУч3Текст");
                SetBinding(doc, Category.GetCategory(doc, BuiltInCategory.OST_Sheets), kpShCU3, true);

                // Колво Уч4
                Definition kpShCU4 = GetParameter_ByGroupAndName(doc, "09 Заполнение штампа", "Ш.КолвоУч4Текст");
                SetBinding(doc, Category.GetCategory(doc, BuiltInCategory.OST_Sheets), kpShCU4, true);

                // Перфикс имени листа
                Definition kpPrViewName = GetParameter_ByGroupAndName(doc, "09 Заполнение штампа", "Ш.ПрефиксИмениЛиста");
                SetBinding(doc, Category.GetCategory(doc, BuiltInCategory.OST_Sheets), kpPrViewName, true);

                // Суффикс имени листа
                Definition kpSuffViewName = GetParameter_ByGroupAndName(doc, "09 Заполнение штампа", "Ш.СуффиксИмениЛиста");
                SetBinding(doc, Category.GetCategory(doc, BuiltInCategory.OST_Sheets), kpSuffViewName, true);

                // Статус листа
                Definition kpViewStatus = GetParameter_ByGroupAndName(doc, "09 Заполнение штампа", "Ш.ШифрСтатусЛиста");
                SetBinding(doc, Category.GetCategory(doc, BuiltInCategory.OST_Sheets), kpViewStatus, true);

                // Статус проекта
                Definition kpPrjStatus = GetParameter_ByGroupAndName(doc, "09 Заполнение штампа", "Ш.СтатусПроекта");
                SetBinding(doc, Category.GetCategory(doc, BuiltInCategory.OST_Sheets), kpPrjStatus, true);

                trans.Commit();
            }
        }

        private Definition GetParameter_ByGroupAndName(Document doc, string paramGroup, string paramName)
        {
            DefinitionFile sharedParamsFile = doc.Application.OpenSharedParameterFile();

            DefinitionGroup sharedParamsGroup = sharedParamsFile.Groups.get_Item(paramGroup);
            Definition definition = sharedParamsGroup.Definitions.get_Item(paramName);
            if (definition != null)
                return definition;
            else
                throw new Exception("Общий параметр с именем '" + paramName + "' не найден в файле параметров.");
        }

        private void SetBinding(Document doc, Category category, Definition definition, bool isInstance)
        {
            // Создаем параметр экземпляра
            CategorySet categorySet = new CategorySet();
            categorySet.Insert(category);

            Binding binding;
            if (isInstance)
                binding = doc.Application.Create.NewInstanceBinding(categorySet);
            else
                binding = doc.Application.Create.NewTypeBinding(categorySet);

            // Добавляем общий параметр к выбранной категории
            BindingMap bindingMap = doc.ParameterBindings;
            bindingMap.Insert(definition, binding, BuiltInParameterGroup.PG_GENERAL);
        }
    }
}
