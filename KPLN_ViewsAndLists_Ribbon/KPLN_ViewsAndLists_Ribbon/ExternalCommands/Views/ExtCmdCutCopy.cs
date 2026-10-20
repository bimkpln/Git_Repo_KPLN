using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_ViewsAndLists_Ribbon.Common.Views;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

namespace KPLN_ViewsAndLists_Ribbon.ExternalCommands.Views
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]

    class ExtCmdCutCopy : IExternalCommand
    {
        internal const string PluginName = "Копировать подрезку";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            if (doc.IsFamilyDocument)
            {
                MessageBox.Show(
                    "Работает только с проектами Revit. Сейчас запущено в семействе",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                return Result.Cancelled;
            }


            // Получаю активное окно 
            View actView = uidoc.ActiveView;
            View actGraphView = uidoc.ActiveGraphicalView;

            // Блок ошибочного запуска
            if (actGraphView == null
                || (actGraphView.ViewType != ViewType.CeilingPlan
                && actGraphView.ViewType != ViewType.EngineeringPlan
                && actGraphView.ViewType != ViewType.FloorPlan))
            {
                MessageBox.Show(
                    "Запускай с активного плана этажа, плана потолков, или плана несущих конструкций (тыкни ЛКМ в рабочей области плана)",
                    "Внимание",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);

                return Result.Cancelled;
            }


            bool actViewCrop = actGraphView.CropBoxActive;
            if (!actViewCrop)
            {
                MessageBox.Show(
                    "На текущем виде подрезка не активна. Копировать то, что не используется - недопустимо. " +
                        "Включи и проанализируй подрезку перед копированием",
                    "Внимание",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);

                return Result.Failed;
            }

            // Собираю планы для анализа
            // 1. Юзер выбрал их в модели
            List<View> selectedViews = new List<View>();
            if (actView.ViewType == ViewType.ProjectBrowser)
            {
                Selection sel = commandData.Application.ActiveUIDocument.Selection;
                ElementId[] selIds = sel.GetElementIds().ToArray();
                foreach (ElementId selId in selIds)
                {
                    Element elem = doc.GetElement(selId);
                    if (elem is View curView)
                        selectedViews.Add(curView);
                }
            }
            // 2. Нужно окно с выбором
            else
            {
                IEnumerable<View> docViewsElem = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>();

                List<ElementEntity> viewEntities = docViewsElem
                    .Where(v =>
                        (v.ViewType == ViewType.CeilingPlan || v.ViewType == ViewType.EngineeringPlan || v.ViewType == ViewType.FloorPlan)
                        && !v.IsTemplate)
                    .Select(el => new ElementEntity(el)).ToList();


                // Формирование окна и осн. процесс
                ElementMultiPick mainForm = new ElementMultiPick(null, viewEntities, "Выбери планы для подрезки");
                if (mainForm.ShowDialog() == true)
                {
                    // Отлов ошибки юзера
                    if (mainForm.SelectedElements.Count == 0)
                    {
                        MessageBox.Show(
                            "Вы не выбрали, куда копировать",
                            "Внимание",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);

                        return Result.Cancelled;
                    }

                    selectedViews = mainForm.SelectedElements.Select(ent => ent.Element).Where(el => el is View).Cast<View>().ToList();
                }
                else
                    return Result.Cancelled;
            }

            if (!selectedViews.Any())
            {
                MessageBox.Show(
                    "Отправь разработчику: Не удалось выловить ошибки ввода",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                return Result.Cancelled;
            }


            // Эталон и неподходящие виды остаются в отчёте, но не изменяются.
            var copied = new List<string>();
            var skipped = new List<string>();
            var failed = new List<string>();
            foreach (View currentView in selectedViews)
            {
                string viewLabel = $"«{currentView.Name}» (ID: {currentView.Id})";
                if (currentView.Id.Equals(actGraphView.Id))
                {
                    skipped.Add(viewLabel + " — пропущен, т.к. это эталон (вид-донор).");
                    continue;
                }

                if (currentView.IsTemplate
                    || (currentView.ViewType != ViewType.CeilingPlan
                    && currentView.ViewType != ViewType.EngineeringPlan
                    && currentView.ViewType != ViewType.FloorPlan))
                {
                    skipped.Add(viewLabel + " — шаблон или неподдерживаемый тип вида; требуется план.");
                    continue;
                }


                // Ошибка одного плана не отменяет уже скопированные подрезки.
                using (Transaction trans = new Transaction(doc, "KPLN: Копировать подрезку"))
                {
                    try
                    {
                        trans.Start();
                        FailureHandlingOptions options = trans.GetFailureHandlingOptions();
                        options.SetClearAfterRollback(true);
                        var failureProcessor = new CropCopyFailures();
                        options.SetFailuresPreprocessor(failureProcessor);
                        trans.SetFailureHandlingOptions(options);

                        ViewCropCopyService.Copy(actGraphView, currentView);
                        if (trans.Commit() == TransactionStatus.Committed)
                            copied.Add(viewLabel);
                        else
                            failed.Add(viewLabel + " — Revit отменил изменения. " + failureProcessor.ErrorText);
                    }
                    catch (Exception ex)
                    {
                        if (trans.GetStatus() == TransactionStatus.Started)
                            trans.RollBack();

                        failed.Add(viewLabel + " — " + ex.Message);
                    }
                }
            }


            string summary = $"Выбрано: {selectedViews.Count} | Скопировано: {copied.Count}"
                + $" | Пропущено: {skipped.Count} | Ошибок: {failed.Count}";
            var report = new StringBuilder();
            report.AppendLine($"Проект: {doc.Title}");
            report.AppendLine($"Эталон: «{actGraphView.Name}» (ID: {actGraphView.Id})");
            report.AppendLine("Переносимые настройки: контур подрезки, включение и видимость подрезки,");
            report.AppendLine("включение подрезки аннотаций и её отступы (если поддерживаются обоими видами).");
            AppendReportSection(report, "СКОПИРОВАНО", copied);
            AppendReportSection(report, "ПРОПУЩЕНО", skipped);
            AppendReportSection(report, "ОШИБКИ — изменения этих видов отменены", failed);
            if (failed.Count > 0)
                report.AppendLine("\nПроверьте шаблон вида, область видимости и доступность вида для редактирования.");

            var resultWindow = new OperationReportWindow("KPLN: Итоги копирования подрезки", report.ToString(), summary);
            WindowHandleSearch.MainWindowHandle.SetAsOwner(resultWindow);
            resultWindow.ShowDialog();

            return Result.Succeeded;
        }

        private static void AppendReportSection(StringBuilder report, string title, List<string> rows)
        {
            report.AppendLine();
            report.AppendLine($"{title}: {rows.Count}");
            foreach (string row in rows)
                report.AppendLine("• " + row);
        }



        private sealed class CropCopyFailures : IFailuresPreprocessor
        {
            public string ErrorText { get; private set; } = "";

            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                var errors = failuresAccessor.GetFailureMessages()
                    .Where(f => f.GetSeverity() != FailureSeverity.Warning).ToList();
                if (errors.Count == 0)
                    return FailureProcessingResult.Continue;

                ErrorText = string.Join("; ", errors.Select(f => f.GetDescriptionText()));
                return FailureProcessingResult.ProceedWithRollBack;
            }
        }
    }
}

