using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_ModelChecker_Lib;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Builder
{
    internal class GripBuilder_KR : AbstrGripBuilder
    {
        public GripBuilder_KR(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName, double floorScreedHeight, double downAndTopExtra) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName, floorScreedHeight, downAndTopExtra)
        {
        }

        public override void Prepare()
        {
            // Таска на подготовку солидов секций/этажей
            Task sectSolidPrepareTask = Task.Run(() =>
            {
                SectDataSolids = LevelAndGridSolid.PrepareSolids(Doc, SectionParamName, FloorScreedHeight);
            });

            // Таска на подготовку элементов на основе (ByHost)
            Task elemsByHostPrepareTask = Task.Run(() =>
            {
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
            });

            // Таска на подготовку элементов под уровнем (ElemsUnderLevel)
            Task elemsUnderLevelPrepareTask = Task.Run(() =>
            {
                // Категория "Стены" монолит под уровнем
                ElemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                    .OfClass(typeof(Wall))
                    .Cast<Wall>()
                    .Where(x => 
                        x.Name.StartsWith("00_") 
                        && (x.Name.ToLower().Contains("перепад") || x.Name.ToLower().Contains("балк")))
                    .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

                // Категория "Перекрытия" монолит под уровнем
                ElemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                    .OfClass(typeof(Floor))
                    .Cast<Floor>()
                    .Where(x => 
                        x.Name.StartsWith("00_") 
                        && (x.Name.ToLower().Contains("площадка") || x.Name.ToLower().Contains("фундамент") || x.Name.ToLower().Contains("пандус")))
                    .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

                // Семейства "Колоны" под уровнем
                ElemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .Cast<FamilyInstance>()
                    .Where(x => x.SuperComponent == null)
                    .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

                // Категория "Кровля" под уровнем
                ElemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                    .OfClass(typeof(RoofBase))
                    .Cast<RoofBase>()
                    .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

                // Семейства "Обобщенная модель" под уровнем
                ElemsUnderLevel.AddRange(new FilteredElementCollector(Doc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .Cast<FamilyInstance>()
                    .Where(x => x.SuperComponent == null)
                    .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));
            });

            // Категория "Стены" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(x => 
                    !x.Name.StartsWith("00_") 
                    || (x.Name.StartsWith("00_") && (!x.Name.ToLower().Contains("перепад") || !x.Name.ToLower().Contains("балк"))))
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            // Категория "Перекрытия" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Floor))
                .Cast<Floor>()
                .Where(x =>
                    !x.Name.StartsWith("00_")
                    || !(x.Name.StartsWith("00_") && (x.Name.ToLower().Contains("площадка") || x.Name.ToLower().Contains("фундамент") || x.Name.ToLower().Contains("пандус"))))
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            // Семейства "Обобщенные модели" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(i => !i.Symbol.FamilyName.StartsWith("22") && i.Symbol.FamilyName.StartsWith("2") && i.SuperComponent == null)
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            // Семейства "Окна" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(i => i.Symbol.FamilyName.StartsWith("23") && i.SuperComponent == null)
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            // Категория "Лестницы" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
               .OfClass(typeof(StairsRun))
               .Cast<StairsRun>()
               .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
               .OfClass(typeof(StairsLanding))
               .Cast<StairsLanding>()
               .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            // Семейства "Колоны" над уровнем
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .Cast<FamilyInstance>()
                .Where(x => x.SuperComponent == null)
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                .Cast<FamilyInstance>()
                .Where(x => x.SuperComponent == null)
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            Task.WaitAll(sectSolidPrepareTask, elemsByHostPrepareTask, elemsUnderLevelPrepareTask);
        }
    }
}
