using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_ModelChecker_Lib.Services.GripGeom.Core;
using System.Linq;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Builder
{
    internal class GripBuilder_AR : AbstrGripBuilder
    {
        public GripBuilder_AR(Document doc, string docMainTitle, string levelParamName, string sectionParamName) : base(doc, docMainTitle, levelParamName, sectionParamName)
        {
        }

        public GripBuilder_AR(Document doc, string docMainTitle, string levelParamName, string sectionParamName, string corpsParamName) : base(doc, docMainTitle, levelParamName, sectionParamName, corpsParamName)
        {
        }

        public override void Prepare()
        {
            // Подготовка в основном потоке Revit: солидов секций/этажей
            SectDataSolids = LevelAndSectionSolid.PrepareSolids(Doc, !string.IsNullOrEmpty(CorpsParamName));

            // Подготовка в основном потоке Revit: элементов на основе (ByHost)
            // Семейства "Панели витража"
            ElemsByHost.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_CurtainWallPanels)
                .Cast<FamilyInstance>()
                .Select(e => new InstanceElemData(e)));

            // Семейства "Импосты витража"
            ElemsByHost.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_CurtainWallMullions)
                .Cast<FamilyInstance>()
                .Select(e => new InstanceElemData(e)));

            // Категория "Стены"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(x => !x.Name.StartsWith("00_") || !x.Name.Contains("КЖ_Монолит"))
                .Select(e => new InstanceGeomData(e)));

            // Категория "Перекрытия"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Floor))
                .Cast<Floor>()
                .Where(x => !x.Name.StartsWith("00_") || !x.Name.Contains("КЖ_Монолит"))
                .Select(e => new InstanceGeomData(e)));

            // Категория "Кровля"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfCategory(BuiltInCategory.OST_Roofs)
                .WhereElementIsNotElementType()
                .Select(e => new InstanceGeomData(e)));

            // Категория "Потолки"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Ceiling))
                .Cast<Ceiling>()
                .Select(e => new InstanceGeomData(e)));

            // Семейства "Окна"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Windows)
                .Cast<FamilyInstance>()
                .Select(e => new InstanceGeomData(e)));

            // Семейства "Двери"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Doors)
                .Cast<FamilyInstance>()
                .Select(e => new InstanceGeomData(e)));

            // Семейства "Парковка"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Parking)
                .Cast<FamilyInstance>()
                .Select(e => new InstanceGeomData(e)));

            // Семейства "Лестничные марши"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Railing))
                .Cast<Railing>()
                .Select(e => new InstanceGeomData(e)));

            // Семейства "Оборудование"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                .Cast<FamilyInstance>()
                //.Where(x => !x.Symbol.FamilyName.StartsWith("199_") && !x.Symbol.FamilyName.Equals("ASML_АР_Шахта"))
                .Select(e => new InstanceGeomData(e)));

            // Семейства "Обощенные модели"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Cast<FamilyInstance>()
                .Where(x => 
                    !x.Symbol.FamilyName.StartsWith("ClashPoint")
                    && !x.Symbol.FamilyName.StartsWith("500_"))
                .Select(e =>    new InstanceGeomData(e)));

            // Семейства "Ограждения"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StairsRailing)
                .Cast<FamilyInstance>()
                .Select(e => new InstanceGeomData(e)));

            // Семейства "Сантехнические приборы"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_PlumbingFixtures)
                .Cast<FamilyInstance>()
                // Только сантехника с геометрией
                .Where(x => x.get_BoundingBox(null).Max.Z > 0 && x.get_BoundingBox(null).Min.Z > 0)
                .Select(e =>    new InstanceGeomData(e)));

            // Семейства "Мебель"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Furniture)
                .Cast<FamilyInstance>()
                // Только мебель с геометрией
                .Where(x => x.get_BoundingBox(null).Max.Z > 0 && x.get_BoundingBox(null).Min.Z > 0)
                .Select(e => new InstanceGeomData(e)));

            // Семейства "Каркас несущий (перемычки)"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .Cast<FamilyInstance>()
                .Where(x => !x.Symbol.FamilyName.StartsWith("199_"))
                .Select(e => new InstanceGeomData(e)));

        }
    }
}