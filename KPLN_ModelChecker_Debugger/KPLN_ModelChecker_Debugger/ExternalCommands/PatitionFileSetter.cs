using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using KPLN_ModelChecker_Lib.Services.GripGeom;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_ModelChecker_Debugger.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal sealed class PatitionFileSetter : IExternalCommand
    {
        private static readonly BuiltInCategory _boxBIC = BuiltInCategory.OST_Mass;

        private static readonly string _fopPath = @"X:\BIM\4_ФОП\02_Для плагинов\КП_Плагины_Общий.txt";
        private static readonly string _upLvlParamName = "ПЗ_Верхний уровень";
        private static readonly string _downLvlParamName = "ПЗ_Нижний уровень";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            // Остортированная коллекция уровней
            List<Level> orderedLvls = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsNotElementType()
                .Cast<Level>()
                .OrderBy(lvl => lvl.Elevation)
                .ToList();

            // Коллекция спец. семейств контура здания
            FamilyInstance[] specialBoxes_Parents = new FilteredElementCollector(doc)
                .OfCategory(_boxBIC)
                .WhereElementIsNotElementType()
                .Where(el => el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("001_Разделитель захваток"))
                .Cast<FamilyInstance>()
                .ToArray();

            // Проверяю/добавляю параметры (можно по родителю, они одной категории со вложенными)
            if (specialBoxes_Parents.Any(fi => fi.LookupParameter(_upLvlParamName) == null || fi.LookupParameter(_downLvlParamName) == null))
                AddParamFromFOP(uiapp);

            // Проверка и установка данных
            List<string> upInfinityBoxesIdColl = new List<string>();
            List<string> downInfinityBoxesIdColl = new List<string>();
            using (Transaction t = new Transaction(doc, "KPLN_РФ: Проверка и установка данных"))
            {
                t.Start();

                foreach (FamilyInstance boxParent in specialBoxes_Parents)
                {
                    // Вложенные боксы в спец. семейства
                    Element[] subBoxes = boxParent
                        .GetSubComponentIds()
                        .Select(id => doc.GetElement(id))
                        .ToArray();
                    foreach (Element subBox in subBoxes)
                    {
                        // Т.к. боксы квадратные - достаточно облегченных BoundingBoxXYZ, а не Solid
                        BoundingBoxXYZ bbox = subBox.get_BoundingBox(null);
                        double minZ = Math.Round(bbox.Min.Z, 3);
                        double maxZ = Math.Round(bbox.Max.Z, 3);


                        // Установка значений параметров
                        Parameter upLvlParam = subBox.LookupParameter(_upLvlParamName);
                        Parameter downLvlParam = subBox.LookupParameter(_downLvlParamName);


                        // Поиск уровней по отметкам
                        Level upLevelFromDoc = LevelWorker.BinaryFindExactLevel(orderedLvls, maxZ);
                        Level downLevelFromDoc = LevelWorker.BinaryFindExactLevel(orderedLvls, minZ);


                        if (upLevelFromDoc == null)
                        {
                            upLvlParam.Set("+∞");
                            upInfinityBoxesIdColl.Add(subBox.Id.ToString());
                        }
                        else
#if Debug2020 || Revit2020
                            upLvlParam.Set(Math.Round(UnitUtils.ConvertFromInternalUnits(upLevelFromDoc.Elevation, DisplayUnitType.DUT_MILLIMETERS), 3).ToString());
#else
                            upLvlParam.Set(Math.Round(UnitUtils.ConvertFromInternalUnits(upLevelFromDoc.Elevation, SpecTypeId.Length), 3).ToString());
#endif


                        if (downLevelFromDoc == null)
                        {
                            downLvlParam.Set("-∞");
                            downInfinityBoxesIdColl.Add(subBox.Id.ToString());
                        }
                        else
#if Debug2020 || Revit2020
                            downLvlParam.Set(Math.Round(UnitUtils.ConvertFromInternalUnits(downLevelFromDoc.Elevation, DisplayUnitType.DUT_MILLIMETERS), 3).ToString());
#else
                            downLvlParam.Set(Math.Round(UnitUtils.ConvertFromInternalUnits(downLevelFromDoc.Elevation, SpecTypeId.Length), 3).ToString());
#endif
                    }
                }


                t.Commit();
            }

            string upElemIds = "Таких не обнаружено, и это - ошибка";
            string downElemIds = "Таких не обнаружено, и это - не обязательно";
            if (upInfinityBoxesIdColl.Any())
                upElemIds = string.Join(",", upInfinityBoxesIdColl);

            if (downInfinityBoxesIdColl.Any())
                downElemIds = string.Join(",", downInfinityBoxesIdColl);

            Print($"Нижние боксы (id): {downElemIds}",
                MessageType.Warning);
            Print($"Верхние боксы (id): {upElemIds}",
                MessageType.Warning);
            Print($"Проверь боксы верхних и нижних этажей, у них не вписаны граничные уровни. Такими могут быть ТОЛЬКО уровни основания, или кровель:",
                MessageType.Warning);

            Print($"Проверка/параметризация завершена успешно! Детали см. ниже:",
                MessageType.Success);

            return Result.Succeeded;
        }

        /// <summary>
        /// Добавить параметры для эл-в модели
        /// </summary>
        private static void AddParamFromFOP(UIApplication uiapp)
        {
            Application app = uiapp.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            app.SharedParametersFilename = _fopPath;
            DefinitionFile fopDefFile = app.OpenSharedParameterFile();

            List<Definition> defsToAdd = new List<Definition>();
            DefinitionGroups fopGroups = fopDefFile.Groups;
            foreach (DefinitionGroup defGroup in fopGroups)
            {
                if (defGroup.Name == "Общие")
                {
                    foreach (Definition definition in defGroup.Definitions)
                    {
                        if (definition.Name == _upLvlParamName || definition.Name == _downLvlParamName)
                            defsToAdd.Add(definition);
                    }
                }
            }


            BindingMap bindingMap = doc.ParameterBindings;
            using (Transaction t = new Transaction(doc, "KPLN: Добавить параметры"))
            {
                t.Start();

                var catSet = app.Create.NewCategorySet();
                catSet.Insert(doc.Settings.Categories.get_Item(_boxBIC));

                foreach (Definition def in defsToAdd)
                {
                    var newBind = app.Create.NewInstanceBinding(catSet);
                    if (!(bindingMap.Insert(def, newBind, BuiltInParameterGroup.PG_TEXT)))
                        bindingMap.ReInsert(def, newBind, BuiltInParameterGroup.PG_TEXT);
                }

                t.Commit();
            }


            bindingMap = doc.ParameterBindings;
            using (Transaction t = new Transaction(doc, "KPLN: Установить параметры"))
            {
                t.Start();

                var it = bindingMap.ForwardIterator();
                it.Reset();

                while (it.MoveNext())
                {
                    if (it.Key is InternalDefinition intDef && (intDef.Name == _upLvlParamName || intDef.Name == _downLvlParamName))
                        intDef.SetAllowVaryBetweenGroups(doc, true);
                }

                t.Commit();
            }
        }
    }
}