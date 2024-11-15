using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.LevelAndGridBoxUtil;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Builder
{
    internal class GripBuilder_IOS : AbstrGripBuilder
    {
        public GripBuilder_IOS(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName, double floorScreedHeight, double downAndTopExtra) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName, floorScreedHeight, downAndTopExtra)
        {
        }

        public override void Prepare()
        {
            // Таска на подготовку солидов секций/этажей
            Task sectSolidPrepareTask = Task.Run(() =>
            {
                SectDataSolids = LevelAndGridSolid.PrepareSolids(Doc, SectionParamName, LevelParamName,
                    FloorScreedHeight, DownAndTopExtra);
            });

            List<BuiltInCategory> userCat = null;
            List<BuiltInCategory> revitCat = null;

            // Делю на ЭОМ СС
            if (Doc.Title.ToUpper().Contains("ЭОМ")
                || Doc.Title.ToUpper().Contains("_EOM")
                || Doc.Title.ToUpper().Contains("_СС")
                || Doc.Title.ToUpper().Contains("_CC")
                || Doc.Title.ToUpper().Contains("_АВ")
                || Doc.Title.ToUpper().Contains("_AV"))
            {
                // Категории пользовательских семейств, используемые в проектах ИОС
                userCat = new List<BuiltInCategory>()
                {
                    BuiltInCategory.OST_MechanicalEquipment,
                    BuiltInCategory.OST_ElectricalEquipment,
                    BuiltInCategory.OST_CableTrayFitting,
                    BuiltInCategory.OST_ConduitFitting,
                    BuiltInCategory.OST_LightingFixtures,
                    BuiltInCategory.OST_ElectricalFixtures,
                    BuiltInCategory.OST_DataDevices,
                    BuiltInCategory.OST_LightingDevices,
                    BuiltInCategory.OST_NurseCallDevices,
                    BuiltInCategory.OST_SecurityDevices,
                    BuiltInCategory.OST_FireAlarmDevices,
                    BuiltInCategory.OST_GenericModel,
                    //Огнезащитное покрытие
                    BuiltInCategory.OST_DuctFitting
                };

                // Категории системных семейств, используемые в проектах ИОС
                revitCat = new List<BuiltInCategory>()
                {
                    BuiltInCategory.OST_CableTray,
                    BuiltInCategory.OST_Conduit,
                    //Огнезащитное покрытие
                    BuiltInCategory.OST_DuctCurves
                };
            }
            // Делю на ОВ ВК
            else
            {
                // Категории пользовательских семейств, используемые в проектах ИОС
                userCat = new List<BuiltInCategory>()
                {
                    BuiltInCategory.OST_MechanicalEquipment,
                    BuiltInCategory.OST_DuctAccessory,
                    BuiltInCategory.OST_DuctFitting,
                    BuiltInCategory.OST_DuctTerminal,
                    BuiltInCategory.OST_PipeAccessory,
                    BuiltInCategory.OST_PipeFitting,
                    BuiltInCategory.OST_Sprinklers,
                    BuiltInCategory.OST_PlumbingFixtures,
                    BuiltInCategory.OST_GenericModel
                };

                // Категории системных семейств, используемые в проектах ИОС
                revitCat = new List<BuiltInCategory>()
                {
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_FlexDuctCurves,
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_FlexPipeCurves,
                };
            }

            // Таска на подготовку элементов на основе (ByHost)
            Task elemsByHostPrepareTask = Task.Run(() =>
            {
                ElemsByHost.AddRange(new FilteredElementCollector(Doc)
                    .OfCategory(BuiltInCategory.OST_DuctInsulations)
                    .WhereElementIsNotElementType()
                    .Select(e => new InstanceElemData(e)));
                ElemsByHost.AddRange(new FilteredElementCollector(Doc)
                    .OfCategory(BuiltInCategory.OST_PipeInsulations)
                    .WhereElementIsNotElementType()
                    .Select(e => new InstanceElemData(e)));
            });

            foreach (BuiltInCategory bic in revitCat)
            {
                ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));
            }
            Task.WaitAll(elemsByHostPrepareTask);

            foreach (BuiltInCategory bic in userCat)
            {
                IEnumerable<FamilyInstance> famInst = new FilteredElementCollector(Doc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(bic)
                    .Cast<FamilyInstance>()
                    .Where(x =>
                        !x.Symbol.FamilyName.StartsWith("ClashPoint")
                        && !x.Symbol.FamilyName.StartsWith("500_")
                        && !x.Symbol.FamilyName.StartsWith("501_")
                        && !x.Symbol.FamilyName.StartsWith("502_")
                        && !x.Symbol.FamilyName.StartsWith("503_"));

                ElemsOnLevel.AddRange(famInst
                    .Where(x => x.SuperComponent == null)
                    .Select(e => new InstanceGeomData(e).SetCurrentSolidColl().SetCurrentBBoxColl()));
                ElemsByHost.AddRange(famInst
                    .Where(x => x.SuperComponent != null)
                    .Select(e => new InstanceElemData(e)));
            }

            Task.WaitAll(sectSolidPrepareTask);
        }
    }
}