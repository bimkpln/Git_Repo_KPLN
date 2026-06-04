using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using KPLN_Tools.Common.TagsHelpers;
using KPLN_Tools.Forms;
using KPLN_Tools.Forms.Models;
using System;
using System.Linq;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class ExtCmd_AUPT_TagPlacer : IExternalCommand
    {
        internal const string PluginName = "АУПТ: Автомаркировка";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = doc.ActiveView;

            if (activeView == null || !(activeView is ViewPlan))
            {
                TaskDialog.Show(
                    PluginName,
                    "Активный вид должен быть планом этажа.");

                return Result.Failed;
            }

            AUPT_TagPlacerForm mainForm = new AUPT_TagPlacerForm(doc);
            if (!(bool)mainForm.ShowDialog())
                return Result.Cancelled;

            AUPTTagPlacerM placerM = mainForm.AUPTTagPlacerViewModel.AUPTTagPlacerModel;

            // 1. Находим марку семейства ASML_ВК_Марка_Труба
            if (placerM.SelectedTagType == null)
            {
                TaskDialog.Show(
                    PluginName,
                    $"Не найдено семейство марки '{placerM.SelectedTagTypeName}'. " +
                            "Загрузите его в проект перед запуском.");

                return Result.Failed;
            }


            // 2. Собираем трубы АУПТ на активном виде
            Pipe[] auptHeapPipes = CollectAuptPipes(doc, activeView, placerM);
            if (auptHeapPipes.Length == 0)
            {
                TaskDialog.Show(PluginName, "На активном виде не найдено труб системы АУПТ. " +
                    "Проверь вид, либо настройки фильтрации окна плагина");

                return Result.Cancelled;
            }


            // 3. Уточняю коллекцию труб по выбору пользователя
            Pipe[] auptPipes;
            if (placerM.IngoreMainPipe)
            {
                // Определяем диаметр магистрали (максимальный среди труб АУПТ)
                double mainDiameter = auptHeapPipes.Max(p => GetPipeDiameter(p));

                // Ответвления — трубы с диаметром меньше магистрали
                auptPipes = auptHeapPipes
                    .Where(p => GetPipeDiameter(p) < mainDiameter - 1e-6)
                    .ToArray();

                // Проверка на ошибки поиска
                if (auptPipes.Length == 0)
                {
#if Debug2020 || Revit2020
                    string descr = $"Не найдено ответвлений. Магистраль Ø{UnitUtils.ConvertFromInternalUnits(mainDiameter, DisplayUnitType.DUT_MILLIMETERS):F0} мм — единственный диаметр.";
#else
                    string descr = $"Не найдено ответвлений. Магистраль Ø{UnitUtils.ConvertFromInternalUnits(mainDiameter, UnitTypeId.Millimeters):F0} мм — единственный диаметр.";
#endif
                    TaskDialog.Show(PluginName, descr);

                    return Result.Cancelled;
                }
            }
            else
                auptPipes = auptHeapPipes;


            // 5. Запускаю маркировку
            using (Transaction t = new Transaction(doc, PluginName))
            {
                t.Start();

                if (!placerM.SelectedTagType.IsActive)
                {
                    placerM.SelectedTagType.Activate();
                    doc.Regenerate();
                }

                var placer = new TagPlacer(doc, activeView, placerM);

                placer.PlaceForBranches(auptPipes);

                t.Commit();
            }

            return Result.Succeeded;
        }

        private Pipe[] CollectAuptPipes(Document doc, View view, AUPTTagPlacerM placerM) =>
            new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .Where(p => IsAuptPipe(p, placerM))
                .ToArray();

        private bool IsAuptPipe(Pipe pipe, AUPTTagPlacerM placerM)
        {
            var systemTypeParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
            if (systemTypeParam != null)
            {
                string systemTypeName = systemTypeParam.AsValueString() ?? string.Empty;
                if (systemTypeName.IndexOf(placerM.AUPTSystemTypeNameMainPart, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }

            return false;
        }

        private static double GetPipeDiameter(Pipe pipe)
        {
            // Внешний диаметр; если нет — номинальный
            var p = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
                 ?? pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
            return p?.AsDouble() ?? 0.0;
        }
    }
}
