using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using KPLN_Tools.Common.TagsHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class ExtCmd_AUPT_TagPlacer : IExternalCommand
    {
        internal const string PluginName = "АУПТ: Автомаркировка";

        private const string TAG_FAMILY_NAME = "ASML_ВК_Марка_Труба";
        private const string SYSTEM_KEYWORD = "АУПТ";

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

            try
            {
                // 1. Находим марку семейства ASML_ВК_Марка_Труба
                FamilySymbol tagSymbol = FindTagSymbol(doc);
                if (tagSymbol == null)
                {
                    TaskDialog.Show(
                        PluginName, 
                        $"Не найдено семейство марки '{TAG_FAMILY_NAME}'. " +
                              "Загрузите его в проект перед запуском.");
                    
                    return Result.Failed;
                }


                // 2. Собираем трубы АУПТ на активном виде
                List<Pipe> auptPipes = CollectAuptPipes(doc, activeView);
                if (auptPipes.Count == 0)
                {
                    TaskDialog.Show(PluginName, "На активном виде не найдено труб системы АУПТ.");
                    
                    return Result.Cancelled;
                }


                // 3. Определяем диаметр магистрали (максимальный среди труб АУПТ)
                double mainDiameter = auptPipes.Max(p => GetPipeDiameter(p));


                // 4. Ответвления — трубы с диаметром меньше магистрали
                List<Pipe> branchPipes = auptPipes
                    .Where(p => GetPipeDiameter(p) < mainDiameter - 1e-6)
                    .ToList();

                if (branchPipes.Count == 0)
                {
#if Debug2020 || Revit2020
                    string descr = $"Не найдено ответвлений. Магистраль Ø{UnitUtils.ConvertFromInternalUnits(mainDiameter, DisplayUnitType.DUT_MILLIMETERS):F0} мм — единственный диаметр.";
#else
                    string descr = $"Не найдено ответвлений. Магистраль Ø{UnitUtils.ConvertFromInternalUnits(mainDiameter, UnitTypeId.Millimeters):F0} мм — единственный диаметр.";
#endif
                    TaskDialog.Show(PluginName, descr);
                    
                    return Result.Cancelled;
                }


                // 5. Запускаю маркировку
                int placedCount;
                int skippedCount;
                using (Transaction t = new Transaction(doc, PluginName))
                {
                    t.Start();

                    if (!tagSymbol.IsActive)
                    {
                        tagSymbol.Activate();
                        doc.Regenerate();
                    }

                    var resolver = new AnnotationCollisionResolver(doc, activeView);
                    var placer = new TagPlacer(doc, activeView, tagSymbol, resolver);

                    placer.PlaceForBranches(branchPipes, out placedCount, out skippedCount);

                    t.Commit();
                }

                //TaskDialog.Show("ASML",
                //    $"Готово.\n" +
                //    $"Магистраль Ø{UnitUtils.ConvertFromInternalUnits(mainDiameter, DisplayUnitType.DUT_MILLIMETERS):F0} мм\n" +
                //    $"Ответвлений (сегментов): {branchPipes.Count}\n" +
                //    $"Марок размещено: {placedCount}\n" +
                //    $"Пропущено (нет места): {skippedCount}");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message + "\n\n" + ex.StackTrace;
                return Result.Failed;
            }
        }

        private FamilySymbol FindTagSymbol(Document doc)
        {
            // Ищем символ марки трубы (категория Pipe Tags) с нужным именем семейства
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_PipeTags)
                .Cast<FamilySymbol>()
                .FirstOrDefault(s => s.FamilyName.Equals(TAG_FAMILY_NAME, StringComparison.OrdinalIgnoreCase) && s.Name.Equals("Диаметр"));
        }

        private List<Pipe> CollectAuptPipes(Document doc, View view)
        {
            return new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .Where(IsAuptPipe)
                .ToList();
        }

        private bool IsAuptPipe(Pipe pipe)
        {
            var systemTypeParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
            if (systemTypeParam != null)
            {
                string systemTypeName = systemTypeParam.AsValueString() ?? string.Empty;
                if (systemTypeName.IndexOf(SYSTEM_KEYWORD, StringComparison.OrdinalIgnoreCase) >= 0)
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
