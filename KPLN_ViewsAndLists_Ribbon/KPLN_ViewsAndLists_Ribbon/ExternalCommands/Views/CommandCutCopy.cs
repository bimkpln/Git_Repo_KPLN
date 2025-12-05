using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ViewsAndLists_Ribbon.ExternalCommands.Views
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]

    class CommandCutCopy : IExternalCommand
    {
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
                    TaskDialog crBoxDialog = new TaskDialog("Ошибка")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconError,
                        MainContent = "Твой текущий вид не содержит подрезки. Копировать нечего"
                    };
                    crBoxDialog.Show();

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

                return Result.Succeeded;
            }

            TaskDialog taskDialog = new TaskDialog("Ошибка")
            {
                MainIcon = TaskDialogIcon.TaskDialogIconError,
                MainContent = "Запускай с активного плана этажа, плана потолков, или плана несущих конструкций"
            };
            taskDialog.Show();

            return Result.Failed;
        }
    }
}
