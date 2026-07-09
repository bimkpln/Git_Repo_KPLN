using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.Services;
using KPLN_Library_PluginActivityWorker;
using KPLN_ViewsAndLists_Ribbon.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace KPLN_ViewsAndLists_Ribbon.ExternalCommands.Views
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    class ExtCmdDeleteUnusedViews : IExternalCommand
    {
        internal const string PluginName = "Удаление неразмещённых видов";

        private static readonly string[] _viewNameExc = new string[1] { "Navisworks" };

        private static readonly string[] _scheduleNameExc = new string[1] { "Navisworks" };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            if (doc.IsFamilyDocument)
            {
                MessageBox.Show(
                    $"KPLN: {PluginName}",
                    "Плагин работает только в проектах Revit. Сейчас открыт файл семейства.");

                return Result.Cancelled;
            }

            // Создаю форму
            DeleteUnusedViewsForm mainForm = new DeleteUnusedViewsForm(uiapp, GetUnusedDocViews(doc));
            WindowHandleSearch.MainWindowHandle.SetAsOwner(mainForm);

            mainForm.Show();

            // Счетчик факта запуска
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(PluginName, ModuleData.ModuleName).ConfigureAwait(false);

            return Result.Succeeded;
        }

        private static View[] GetUnusedDocViews(Document doc)
        {
            HashSet<ElementId> placedViewIds = GetPlacedViewIds(doc);

            // Коллекция видов, у которых есть видовые экраны и она размещены на листах
            List<View> views = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Views)
                .WhereElementIsNotElementType()
                .Cast<View>()
                .Where(v => !v.IsTemplate
                    && !placedViewIds.Contains(v.Id)
                    && !NameContainsAny(v.Name, _viewNameExc))
                .ToList();

            // Дополняю коллекцией спецификаций
            views.AddRange(GetUnusedSchedules(doc));

            return views.ToArray();
        }

        /// <summary>
        /// Получить список ID всех размещённых на листах видов (не попадут спецификации)
        /// </summary>
        private static HashSet<ElementId> GetPlacedViewIds(Document doc)
        {
            HashSet<ElementId> placedViewIds = new HashSet<ElementId>();

            List<ViewSheet> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(s => !s.IsPlaceholder)
                .ToList();

            foreach (ViewSheet sheet in sheets)
            {
                foreach (ElementId viewId in sheet.GetAllPlacedViews())
                {
                    placedViewIds.Add(viewId);
                }
            }

            return placedViewIds;
        }

        /// <summary>
        /// Получить список всех спецификаций, которые размещены на листах
        /// </summary>
        private static List<ViewSchedule> GetUnusedSchedules(Document doc)
        {
            HashSet<ElementId> placedScheduleIds = new HashSet<ElementId>();

            List<ScheduleSheetInstance> scheduleSheetInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(ScheduleSheetInstance))
                .Cast<ScheduleSheetInstance>()
                .ToList();

            foreach (ScheduleSheetInstance scheduleSheetInstance in scheduleSheetInstances)
            {
                placedScheduleIds.Add(scheduleSheetInstance.ScheduleId);
            }

            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(s => !s.IsTemplate
                    && !placedScheduleIds.Contains(s.Id)
                    && !IsKeySchedule(s)
                    && !NameContainsAny(s.Name, _scheduleNameExc))
                .ToList();
        }

        private static bool IsKeySchedule(ViewSchedule schedule)
        {
            Parameter scheduleTypeParameter =
                schedule.get_Parameter(BuiltInParameter.SCHEDULE_TYPE_FOR_BROWSER);

            string scheduleTypeName = scheduleTypeParameter == null
                ? string.Empty
                : scheduleTypeParameter.AsValueString();

            if (string.IsNullOrEmpty(scheduleTypeName))
                return false;

            return scheduleTypeName.IndexOf("Ключ", StringComparison.OrdinalIgnoreCase) >= 0
                || scheduleTypeName.IndexOf("Key", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool NameContainsAny(string name, IEnumerable<string> excParts)
        {
            if (string.IsNullOrWhiteSpace(name) || excParts == null)
                return false;

            return excParts.Any(exc =>
                !string.IsNullOrWhiteSpace(exc) &&
                name.IndexOf(exc, StringComparison.OrdinalIgnoreCase) >= 0);
        }
    }
}