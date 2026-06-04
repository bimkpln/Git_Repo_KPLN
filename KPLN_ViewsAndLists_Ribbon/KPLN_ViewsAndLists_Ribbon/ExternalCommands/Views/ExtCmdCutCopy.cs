using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_PluginActivityWorker;
using System.Collections.Generic;
using System.Linq;
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
            string actGraphViewName = actGraphView.Title;

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


            ViewCropRegionShapeManager vcrsManager = actGraphView.GetCropRegionShapeManager();

            // Настройки подрезки вида 
            bool actViewCrop = actGraphView.CropBoxActive;
            bool actViewCropBoxVisible = actGraphView.CropBoxVisible;
            IList<CurveLoop> actViewCropCurves = vcrsManager.GetCropShape();
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

            // Подрезка не может содержать несколько CurveLoop
            CurveLoop actViewCropCurveLoop = actViewCropCurves.FirstOrDefault();

            // Настройки подрезки аннотаций
            int actViewAnnCropParamData = actGraphView.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE).AsInteger();
            double bottomAnnCropOffset = 0;
            double leftAnnCropOffset = 0;
            double rightAnnCropOffset = 0;
            double topAnnCropOffset = 0;
            if (vcrsManager.CanHaveAnnotationCrop)
            {
                bottomAnnCropOffset = vcrsManager.BottomAnnotationCropOffset;
                leftAnnCropOffset = vcrsManager.LeftAnnotationCropOffset;
                rightAnnCropOffset = vcrsManager.RightAnnotationCropOffset;
                topAnnCropOffset = vcrsManager.TopAnnotationCropOffset;
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
                        && !(v.Id.Equals(actGraphView.Id))
                        && !v.IsTemplate)
                    .Select(el => new ElementEntity(el)).ToList();


                // Формирование окна и осн. процесс
                ElementMultiPick mainForm = new ElementMultiPick(null, viewEntities, "Выбери планы для подрезки");
                if ((bool)mainForm.ShowDialog())
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


            // Основной процесс
            using (Transaction trans = new Transaction(doc, "KPLN: Копировать подрезку"))
            {
                trans.Start();

                foreach (View currentView in selectedViews)
                {
                    ViewCropRegionShapeManager currentVCRSManager = currentView.GetCropRegionShapeManager();


                    // Подрезка вида
                    currentView.CropBoxActive = true;
                    currentView.CropBoxVisible = actViewCropBoxVisible;
                    currentVCRSManager.SetCropShape(actViewCropCurveLoop);

                    // Подрезка аннотаций
                    currentView.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE).Set(actViewAnnCropParamData);
                    if (currentVCRSManager.CanHaveAnnotationCrop)
                    {
                        currentVCRSManager.BottomAnnotationCropOffset = bottomAnnCropOffset;
                        currentVCRSManager.LeftAnnotationCropOffset = leftAnnCropOffset;
                        currentVCRSManager.RightAnnotationCropOffset = rightAnnCropOffset;
                        currentVCRSManager.TopAnnotationCropOffset = topAnnCropOffset;
                    }
                }

                trans.Commit();
            }


            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(PluginName, ModuleData.ModuleName).ConfigureAwait(false);


            MessageBox.Show(
                $"Подрезка успешно скопирована для планов, в количестве {selectedViews.Count} шт.",
                "Итог",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            return Result.Succeeded;
        }
    }
}
