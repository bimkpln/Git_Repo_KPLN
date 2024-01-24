using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_ModelChecker_Lib;
using System.Linq;
using System.Threading.Tasks;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Builder
{
    internal class GripBuilder_AR : AbstrGripBuilder
    {
        public GripBuilder_AR(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName, double floorScreedHeight, double downAndTopExtra) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName, floorScreedHeight, downAndTopExtra)
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
            });

            // Категория "Стены"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(x => !x.Name.StartsWith("00_"))
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            // Категория "Перекрытия"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Floor))
                .Cast<Floor>()
                .Where(x => !x.Name.StartsWith("00_"))
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            // Категория "Кровля"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(RoofBase))
                .Cast<RoofBase>()
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            // Категория "Потолки"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Ceiling))
                .Cast<Ceiling>()
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            // Семейства "Окна"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Windows)
                .Cast<FamilyInstance>()
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            // Семейства "Двери"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Doors)
                .Cast<FamilyInstance>()
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            // Семейства "Лестничные марши"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Railing))
                .Cast<Railing>()
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            // Семейства "Оборудование"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                .Cast<FamilyInstance>()
                .Where(x => !x.Symbol.FamilyName.StartsWith("199_"))
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            // Семейства "Обощенные модели"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Cast<FamilyInstance>()
                .Where(x => !x.Symbol.FamilyName.StartsWith("199_"))
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            // Семейства "Каркас несущий (перемычки)"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .Cast<FamilyInstance>()
                .Where(x => !x.Symbol.FamilyName.StartsWith("199_"))
                .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));

            Task.WaitAll(sectSolidPrepareTask, elemsByHostPrepareTask);
        }
    }
}
