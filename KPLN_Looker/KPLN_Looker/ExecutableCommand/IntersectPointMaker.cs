using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace KPLN_Looker.ExecutableCommand
{
    internal class IntersectPointMaker : IExecutableCommand
    {
        private static readonly string _familyName = "ClashPoint_Small";
        private static string _revitVersion;
        private readonly List<XYZ> _intersectPoints;

        public IntersectPointMaker(List<XYZ> intersectPoints)
        {
            _intersectPoints = intersectPoints;
        }

        public Result Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;
            UIDocument uidoc = app.ActiveUIDocument;
            if (string.IsNullOrEmpty(_revitVersion))
                _revitVersion = app.Application.VersionNumber;

            if (uidoc == null)
                return Result.Cancelled;

            using (Transaction trans = new Transaction(doc, "KPLN: Создать точки"))
            {
                trans.Start();

                foreach (XYZ point in _intersectPoints)
                {
                    CreateIntersectFamilyInstance(doc, point);
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Разместить экземпляр семейства пересечения по указанным координатам
        /// </summary>
        private static void CreateIntersectFamilyInstance(Document doc, XYZ point)
        {
            Level level = GetNearestLevel(doc, point.Z) ?? throw new Exception("В проекте отсутсвуют уровни!");

            FamilySymbol intersectFamSymb = GetIntersectFamilySymbol(doc);

            FamilyInstance instance = doc
                .Create
                .NewFamilyInstance(point, intersectFamSymb, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

            doc.Regenerate();
            instance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(point.Z - level.Elevation);
            doc.Regenerate();
        }

        /// <summary>
        /// Получить семейство отображающее коллизию
        /// </summary>
        private static FamilySymbol GetIntersectFamilySymbol(Document doc)
        {
            FamilySymbol[] oldFamSymbOfGM = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Where(el => el is FamilySymbol famSymb && famSymb.FamilyName == _familyName)
                .Cast<FamilySymbol>()
                .ToArray();

            // Если в проекте нет - то грузим
            if (!oldFamSymbOfGM.Any())
            {
                string path = $@"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}\Source\RevitData\{_revitVersion}\{_familyName}.rfa";
                bool result = doc.LoadFamily(path);
                if (!result)
                    throw new Exception("Семейство для метки не найдено! Обратись к разработчику.");

                doc.Regenerate();

                // Повторяем после загрузки семейства
                oldFamSymbOfGM = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .Where(el => el is FamilySymbol famSymb && famSymb.FamilyName == _familyName)
                    .Cast<FamilySymbol>()
                    .ToArray();
            }

            FamilySymbol searchSymbol = oldFamSymbOfGM.FirstOrDefault();
            searchSymbol.Activate();

            return searchSymbol;
        }

        /// <summary>
        ///  Поиск ближайшего подходящего уровня
        /// </summary>
        private static Level GetNearestLevel(Document doc, double elevation)
        {
            Level result = null;

            double resultDistance = 999999;
            foreach (Level lvl in GetLevels(doc))
            {
                double tempDistance = Math.Abs(lvl.Elevation - elevation);
                if (Math.Abs(lvl.Elevation - elevation) < resultDistance)
                {
                    result = lvl;
                    resultDistance = tempDistance;
                }
            }
            return result;
        }

        /// <summary>
        /// Получить коллекцию ВСЕХ уровней проекта
        /// </summary>
        public static Level[] GetLevels(Document doc)
        {
            Level[] instances = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsNotElementType()
                .Cast<Level>()
                .ToArray();

            return instances;
        }

    }
}
