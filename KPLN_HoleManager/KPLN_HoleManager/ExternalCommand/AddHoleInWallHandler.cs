using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Controls;

namespace KPLN_HoleManager.ExternalCommand
{
    public class AddHoleInWallHandler : IExternalEventHandler
    {
        private Document _doc;
        private Element _selectedElement;
        private string _departmentHoleName;
        private string _holeTypeName;
        private string _pluginPath;

        public void SetData(Document doc, Element selectedElement, string departmentHoleName, string holeTypeName, string pluginPath)
        {
            _doc = doc;
            _selectedElement = selectedElement;
            _departmentHoleName = departmentHoleName;
            _holeTypeName = holeTypeName;
            _pluginPath = pluginPath;
        }

        public void Execute(UIApplication app)
        {
            if (_doc == null || _selectedElement == null)
            {
                TaskDialog.Show("Ошибка", "Некорректные входные данные.");
                return;
            }

            string familyFileName = GetFamilyFileName();
            if (string.IsNullOrEmpty(familyFileName))
            {
                TaskDialog.Show("Ошибка", "Не найдено соответствующее семейство.");
                return;
            }

            string familyPath = Path.Combine(_pluginPath, "Families", familyFileName);
            FamilySymbol holeSymbol = LoadFamilySymbol(familyPath);

            if (holeSymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не удалось загрузить семейство.");
                return;
            }

            PlaceFamilyInstance(holeSymbol);
        }

        private string GetFamilyFileName()
        {
            Dictionary<string, string> familyFiles = new Dictionary<string, string>
            {
                { "АР_SquareHole", "199_AR_OSW.rfa" },
                { "АР_RoundHole", "199_AR_ORW.rfa" },
                { "КР_SquareHole", "199_STR_OSW.rfa" },
                { "КР_RoundHole", "199_STR_ORW.rfa" },
                { "ИОС_SquareHole", "501_MEP_TSW.rfa" },
                { "ИОС_RoundHole", "501_MEP_TRW.rfa" }
            };

            string key = $"{_departmentHoleName}_{_holeTypeName}";
            return familyFiles.ContainsKey(key) ? familyFiles[key] : null;
        }

        public string GetName()
        {
            return "AddHoleInWallHandler";
        }

        private FamilySymbol LoadFamilySymbol(string familyPath)
        {
            using (Transaction tx = new Transaction(_doc, "Загрузка семейства"))
            {
                tx.Start();
                Family family = null;

                if (!_doc.LoadFamily(familyPath, out family) || family == null)
                {
                    tx.RollBack();
                    return null;
                }

                FamilySymbol symbol = family.GetFamilySymbolIds()
                    .Select(id => _doc.GetElement(id) as FamilySymbol)
                    .FirstOrDefault();

                tx.Commit();
                return symbol;
            }
        }

        private void PlaceFamilyInstance(FamilySymbol holeSymbol)
        {
            using (Transaction tx = new Transaction(_doc, "Создание отверстия"))
            {
                tx.Start();

                if (!holeSymbol.IsActive)
                    holeSymbol.Activate();

                LocationCurve wallLocation = _selectedElement.Location as LocationCurve;
                if (wallLocation == null)
                {
                    TaskDialog.Show("Ошибка", "Невозможно получить положение стены.");
                    tx.RollBack();
                    return;
                }

                Line wallLine = wallLocation.Curve as Line;
                XYZ wallMidpoint = wallLine.Evaluate(0.5, true);
                XYZ placementPoint = new XYZ(wallMidpoint.X, wallMidpoint.Y, wallMidpoint.Z + 1.0);

                _doc.Create.NewFamilyInstance(placementPoint, holeSymbol, _selectedElement, StructuralType.NonStructural);

                tx.Commit();
            }
        }
    }
}