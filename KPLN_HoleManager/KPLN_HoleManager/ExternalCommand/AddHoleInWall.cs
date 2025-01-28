using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace KPLN_HoleManager.ExternalEventHandlers
{
    public class PlaceHolesEventHandler : IExternalEventHandler
    {
        public void Execute(UIApplication uiApp)
        {
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();

            ElementId selectedId = selectedIds.First();
            Element element = doc.GetElement(selectedId);

            if (element is Wall wall)
            {
                // Пользователь выбирает точку на стене
                XYZ userPoint = null;

                try
                {
                    userPoint = uiDoc.Selection.PickPoint("Выберите точку для размещения отверстия.");
                }
                catch
                {
                    TaskDialog.Show("Отмена", "Действие отменено пользователем.");
                    return;
                }

                using (Transaction trans = new Transaction(doc, "Вставка отверстия"))
                {
                    trans.Start();

                    string pluginFolder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                    string familyPath = Path.Combine(pluginFolder, "Families", "199_AR_ORW.rfa");

                    Family family = LoadFamilyIfNotLoaded(doc, familyPath);

                    if (family == null)
                    {
                        TaskDialog.Show("Ошибка", "Не удалось загрузить семейство 199_AR_ORW.rfa");
                        trans.RollBack();
                        return;
                    }

                    FamilySymbol symbol = GetFirstFamilySymbol(doc, family);
                    if (symbol == null)
                    {
                        TaskDialog.Show("Ошибка", "Не найден доступный тип семейства.");
                        trans.RollBack();
                        return;
                    }

                    if (!symbol.IsActive)
                        symbol.Activate();

                    // Убедимся, что точка принадлежит плоскости стены
                    LocationCurve location = wall.Location as LocationCurve;
                    if (location != null)
                    {
                        Curve curve = location.Curve;
                        if (!IsPointOnCurve(curve, userPoint))
                        {
                            TaskDialog.Show("Ошибка", "Выбранная точка не находится на стене.");
                            trans.RollBack();
                            return;
                        }
                    }

                    // Размещение семейства в выбранной пользователем точке
                    doc.Create.NewFamilyInstance(userPoint, symbol, wall, StructuralType.NonStructural);

                    trans.Commit();
                }
            }
            else
            {
                TaskDialog.Show("Предупреждение", "Выбранный элемент не является стеной.\nПожалуйста, выберите стену.");
            }
        }

        public string GetName()
        {
            return "PlaceHoles EventHandler";
        }

        private Family LoadFamilyIfNotLoaded(Document doc, string familyPath)
        {
            Family family = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .FirstOrDefault(f => f.Name == "199_AR_ORW") as Family;

            if (family == null)
            {
                if (!File.Exists(familyPath))
                {
                    TaskDialog.Show("Ошибка", $"Файл семейства не найден:\n{familyPath}");
                    return null;
                }

                if (!doc.LoadFamily(familyPath, out family))
                {
                    return null;
                }
            }

            return family;
        }

        private FamilySymbol GetFirstFamilySymbol(Document doc, Family family)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                .Cast<FamilySymbol>()
                .FirstOrDefault(s => s.Family.Id == family.Id);
        }

        private bool IsPointOnCurve(Curve curve, XYZ point)
        {
            double tolerance = 0.01; // Допустимая погрешность
            return curve.Project(point).Distance < tolerance;
        }
    }
}