using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_ModelChecker_Lib.Services.GripGeom.Core;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Builder
{
    internal class GripBuilder_KR : AbstrGripBuilder
    {
        public GripBuilder_KR(Document doc, string docMainTitle, string levelParamName, string sectionParamName) : base(doc, docMainTitle, levelParamName, sectionParamName)
        {
        }

        public GripBuilder_KR(Document doc, string docMainTitle, string levelParamName, string sectionParamName, string corpsParamName) : base(doc, docMainTitle, levelParamName, sectionParamName, corpsParamName)
        {
        }

        public override void Prepare()
        {
            // Подготовка в основном потоке Revit: солидов секций/этажей
            SectDataSolids = LevelAndSectionSolid.PrepareSolids(Doc, !string.IsNullOrEmpty(CorpsParamName));

            // Подготовка в основном потоке Revit: элементов на основе (ByHost)
            List<BuiltInCategory> userCat = new List<BuiltInCategory>()
            {
                BuiltInCategory.OST_GenericModel,
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_Rebar,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Floors,
            };

            foreach (BuiltInCategory cat in userCat)
            {
                ElemsByHost.AddRange(new FilteredElementCollector(Doc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(cat)
                    .Cast<FamilyInstance>()
                    .Where(x => x.SuperComponent != null)
                    .Select(e => new InstanceElemData(e)));
            }

            // Подготовка в основном потоке Revit: элементов под уровнем (ElemsUnderLevel)
            // Категория "Стены" монолит под уровнем
            ElemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(x =>
                    x.Name.StartsWith("00_")
                    && (x.Name.ToLower().Contains("перепад") || x.Name.ToLower().Contains("балк")))
                .Select(e => new InstanceGeomData(e)));

            // Категория "Перекрытия" монолит под уровнем
            ElemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Floor))
                .Cast<Floor>()
                .Where(x =>
                    x.Name.StartsWith("00_")
                    && (x.Name.ToLower().Contains("площадка") || x.Name.ToLower().Contains("фундамент") || x.Name.ToLower().Contains("пандус")))
                .Select(e => new InstanceGeomData(e)));

            // Семейства "Колоны" под уровнем
            ElemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .Cast<FamilyInstance>()
                .Where(x => x.SuperComponent == null)
                .Select(e => new InstanceGeomData(e)));

            // Категория "Кровля" под уровнем
            ElemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(RoofBase))
                .Cast<RoofBase>()
                .Select(e => new InstanceGeomData(e)));

            // Семейства "Обобщенная модель" под уровнем
            ElemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Cast<FamilyInstance>()
                .Where(x => x.SuperComponent == null
                        && !x.Symbol.FamilyName.StartsWith("ClashPoint"))
                .Select(e => new InstanceGeomData(e)));

            // Категория "Стены" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Select(e => new InstanceGeomData(e)));

            // Категория "Перекрытия" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Floor))
                .Cast<Floor>()
                .Select(e => new InstanceGeomData(e)));

            // Семейства "Перекрытия" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Floors)
                .Cast<FamilyInstance>()
                .Where(x =>
                    ElemsUnderLevel.Any(ent => !ent.IEDElem.Id.Equals(x.Id)))
                .Select(e => new InstanceGeomData(e)));

            // Семейства "Обобщенные модели" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Cast<FamilyInstance>()
                .Where(x => 
                    ElemsUnderLevel.Any(ent => !ent.IEDElem.Id.Equals(x.Id)))
                .Select(e => new InstanceGeomData(e)));

            // Семейства "Окна" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Windows)
                .Cast<FamilyInstance>()
                .Where(i => i.Symbol.FamilyName.StartsWith("23") && i.SuperComponent == null)
                .Select(e => new InstanceGeomData(e)));

            // Категория "Лестницы" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
               .OfClass(typeof(Stairs))
               .Cast<Stairs>()
               .Select(e => new InstanceGeomData(e)));
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
               .OfClass(typeof(StairsRun))
               .Cast<StairsRun>()
               .Select(e => new InstanceGeomData(e)));
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
               .OfClass(typeof(StairsLanding))
               .Cast<StairsLanding>()
               .Select(e => new InstanceGeomData(e)));

            // Семейства "Колоны" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .Cast<FamilyInstance>()
                .Where(x => x.SuperComponent == null)
                .Select(e => new InstanceGeomData(e)));
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                .Cast<FamilyInstance>()
                .Where(x => x.SuperComponent == null)
                .Select(e => new InstanceGeomData(e)));

        }
    }
}