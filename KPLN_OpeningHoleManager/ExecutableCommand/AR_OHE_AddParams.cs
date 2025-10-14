using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_ModelChecker_Lib;
using KPLN_OpeningHoleManager.Core;
using System;
using System.Collections.Generic;

namespace KPLN_OpeningHoleManager.ExecutableCommand
{
    /// <summary>
    /// Класс по добавлению параметров в семейства отверстий
    /// </summary>
    internal sealed class AR_OHE_AddParams : IExecutableCommand
    {
        private static readonly string _fopPath = @"X:\BIM\4_ФОП\02_Для плагинов\КП_Плагины_Общий.txt";
        private readonly AROpeningHoleEntity _arEntity;

        public AR_OHE_AddParams(AROpeningHoleEntity arEntity)
        {
            _arEntity = arEntity;
        }

        public Result Execute(UIApplication uiapp)
        {
            // Такое возможно при работе плагинов с открываением/сохранением моделей (модель не открылась)
            if (uiapp.ActiveUIDocument == null) return Result.Cancelled;

            Document doc = uiapp.ActiveUIDocument.Document;
            Application app = uiapp.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            if (uidoc == null) return Result.Cancelled;

            if (_arEntity.IEDElem.LookupParameter(AROpeningHoleEntity.AR_OHE_ParamNameCancelOverwrite) != null)
                return Result.Succeeded;

            try
            {
                app.SharedParametersFilename = _fopPath;
                DefinitionFile fopDefFile = app.OpenSharedParameterFile();

                List<Definition> defsToAdd = new List<Definition>();
                DefinitionGroups fopGroups = fopDefFile.Groups;
                foreach (DefinitionGroup defGroup in fopGroups)
                {
                    if (defGroup.Name == "АР_Отверстия")
                    {
                        foreach (Definition definition in defGroup.Definitions)
                        {
                            if (definition.Name == AROpeningHoleEntity.AR_OHE_ParamNameCancelOverwrite)
                                defsToAdd.Add(definition);
                        }
                    }
                }

                BuiltInCategory bic = (BuiltInCategory)_arEntity.IEDElem.Category.Id.IntegerValue;
                BindingMap bindingMap = doc.ParameterBindings;
                using (Transaction t = new Transaction(doc, "KPLN: Добавить параметры"))
                {
                    t.Start();

                    var catSet = app.Create.NewCategorySet();
                    catSet.Insert(doc.Settings.Categories.get_Item(bic));

                    foreach (Definition def in defsToAdd)
                    {
                        var newBind = app.Create.NewInstanceBinding(catSet);
                        if (!(bindingMap.Insert(def, newBind, BuiltInParameterGroup.PG_TEXT)))
                            bindingMap.ReInsert(def, newBind, BuiltInParameterGroup.PG_TEXT);
                    }

                    t.Commit();
                }


                //bindingMap = doc.ParameterBindings;
                //using (Transaction t = new Transaction(doc, "KPLN: Установить параметры"))
                //{
                //    t.Start();

                //    var it = bindingMap.ForwardIterator();
                //    it.Reset();

                //    while (it.MoveNext())
                //    {
                //        if (it.Key is InternalDefinition intDef && (intDef.Name == AROpeningHoleEntity.AR_OHE_ParamNameCancelOverwrite))
                //            intDef.SetAllowVaryBetweenGroups(doc, true);
                //    }

                //    t.Commit();
                //}
            }
            catch (CheckerException ex)
            {
                new TaskDialog("Ошибка")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconWarning,
                    MainInstruction = $"{ex.Message}",
                }.Show();

                return Result.Cancelled;
            }
            catch (Exception ex) { throw ex; }


            return Result.Succeeded;
        }
    }
}
