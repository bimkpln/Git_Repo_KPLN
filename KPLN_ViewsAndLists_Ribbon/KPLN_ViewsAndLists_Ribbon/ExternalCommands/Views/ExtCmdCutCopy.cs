using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
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

            View activeView = doc.ActiveView;
            if (activeView != null 
                && (activeView.ViewType == ViewType.CeilingPlan
                || activeView.ViewType == ViewType.EngineeringPlan
                || activeView.ViewType == ViewType.FloorPlan))
            {
                bool activeViewCrop = activeView.CropBoxActive;
                bool activeViewCropBoxVisible = activeView.CropBoxVisible;
                IList<CurveLoop> activeViewCropCurves= activeView.GetCropRegionShapeManager().GetCropShape();
                if (!activeViewCrop)
                {
                    MessageBox.Show(
                        "Твой текущий вид не содержит подрезки. Копировать нечего",
                        "Внимание",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);

                    return Result.Failed;
                }

                IEnumerable<View> docViewsElem = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>();
                List<ElementEntity> viewEntities = docViewsElem
                    .Where(v => 
                        (v.ViewType == ViewType.CeilingPlan || v.ViewType == ViewType.EngineeringPlan || v.ViewType == ViewType.FloorPlan)
                        && !(v.Id.Equals(activeView.Id))
                        && !v.IsTemplate)
                    .Select(el => new ElementEntity(el)).ToList();

                ElementMultiPick mainForm = new ElementMultiPick(null, viewEntities, "Выбери планы для подрезки");
                mainForm.ShowDialog();

                if (mainForm.SelectedElements.Count > 0)
                {
                    using (Transaction trans = new Transaction(doc, "KPLN: Копировать подрезку"))
                    {
                        trans.Start();

                        foreach (ElementEntity entity in mainForm.SelectedElements)
                        {
                            if (entity.Element is View currentView)
                            {
                                currentView.CropBoxActive = true;
                                currentView.GetCropRegionShapeManager().SetCropShape(activeViewCropCurves[0]);
                                currentView.CropBoxVisible = activeViewCropBoxVisible;
                            }
                            else
                                throw new System.Exception("Отправь разработчику: Ошибка приведения типа из формы в View");
                        }

                        trans.Commit();
                    }
                }


                DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(PluginName, ModuleData.ModuleName).ConfigureAwait(false);

                
                MessageBox.Show(
                    $"Подрезка успешно скопирована для планов, в количестве {mainForm.SelectedElements.Count} шт.",
                    "Итог",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                return Result.Succeeded;
            }

            MessageBox.Show(
                "Запускай с активного плана этажа, плана потолков, или плана несущих конструкций", 
                "Внимание", 
                MessageBoxButton.OK, 
                MessageBoxImage.Warning);

            return Result.Cancelled;
        }
    }
}
