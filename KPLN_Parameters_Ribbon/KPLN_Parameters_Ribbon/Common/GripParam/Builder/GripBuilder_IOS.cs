using Autodesk.Revit.DB;
using KPLN_Parameters_Ribbon.Common.Tools;
using KPLN_Parameters_Ribbon.Forms;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Loader.Output.Output;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Builder
{
    internal class GripBuilder_IOS : AbstrGripBuilder
    {
        public GripBuilder_IOS(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName)
        {
        }

        public GripBuilder_IOS(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName, char splitLevelChar) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName, splitLevelChar)
        {
        }

        public override void Prepare()
        {
            List<BuiltInCategory> userCat = null;
            List<BuiltInCategory> revitCat = null;
            List<BuiltInCategory> revitInsulCat = null;
            List<FamilyInstance> _dirtyElems = new List<FamilyInstance>();

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

                // Категории системных семейств (изоляции), используемые в проектах ИОС
                revitInsulCat = new List<BuiltInCategory>()
                {
                    BuiltInCategory.OST_DuctInsulations,
                    BuiltInCategory.OST_PipeInsulations,
                };

                foreach (BuiltInCategory bic in revitInsulCat)
                {
                    ElemsInsulation.AddRange(new FilteredElementCollector(Doc)
                        .OfCategory(bic)
                        .WhereElementIsNotElementType());
                }
            }

            foreach (BuiltInCategory bic in userCat)
            {
                switch (bic)
                {
                    case BuiltInCategory.OST_MechanicalEquipment:
                        _dirtyElems.AddRange(new FilteredElementCollector(Doc)
                            .OfClass(typeof(FamilyInstance))
                            .OfCategory(bic)
                            .Cast<FamilyInstance>()
                            .Where(x => !x.Symbol.FamilyName.StartsWith("500_"))
                            .Where(x => !x.Symbol.FamilyName.StartsWith("501_"))
                            .Where(x => !x.Symbol.FamilyName.StartsWith("502_"))
                            .Where(x => !x.Symbol.FamilyName.StartsWith("503_")));
                        break;

                    case BuiltInCategory.OST_GenericModel:
                        _dirtyElems.AddRange(new FilteredElementCollector(Doc)
                            .OfClass(typeof(FamilyInstance))
                            .OfCategory(bic)
                            .Cast<FamilyInstance>()
                            .Where(x => !x.Symbol.FamilyName.StartsWith("500_"))
                            .Where(x => !x.Symbol.FamilyName.StartsWith("501_"))
                            .Where(x => !x.Symbol.FamilyName.StartsWith("502_"))
                            .Where(x => !x.Symbol.FamilyName.StartsWith("503_")));
                        break;

                    default:
                        _dirtyElems.AddRange(new FilteredElementCollector(Doc)
                            .OfClass(typeof(FamilyInstance))
                            .OfCategory(bic)
                            .Cast<FamilyInstance>());
                        break;
                }
            }

            ElemsOnLevel.AddRange(_dirtyElems
                .Where(x => x.SuperComponent == null));

            ElemsByHost.AddRange(_dirtyElems
                .Where(x => x.SuperComponent != null));

            foreach (BuiltInCategory bic in revitCat)
            {
                ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType());
            }


        }

        public override bool ExecuteGripParams(Progress_Single pb)
        {
            GripByGeometry gripByGeometry = new GripByGeometry(Doc, LevelParamName, SectionParamName);

            IEnumerable<MyLevel> myLevelColl = gripByGeometry.LevelPrepare();

            Dictionary<string, HashSet<Grid>> gridsDict = gripByGeometry.GridPrepare("КП_О_Секция");

            IEnumerable<MySolid> mySolids = gripByGeometry.SolidsCollectionPrepare(gridsDict, myLevelColl, LevelNumberIndex, SplitLevelChar);

            bool byElem = false;
            if (ElemsOnLevel.Count > 0)
            {

                IEnumerable<Element> notIntersectedElems = gripByGeometry.IntersectWithSolidExcecute(ElemsOnLevel, mySolids, pb);

                IEnumerable<Element> notNearestSolidElems = gripByGeometry.FindNearestSolid(notIntersectedElems, mySolids, pb);

                IEnumerable<Element> notRevalueElems = gripByGeometry.ReValueDuplicates(mySolids, pb);

                if (gripByGeometry.DuplicatesWriteParamElems.Keys.Count() > 0)
                {
                    foreach (Element element in gripByGeometry.DuplicatesWriteParamElems.Keys)
                    {
                        Print($"Проверь вручную элемент с id: {element.Id}", KPLN_Loader.Preferences.MessageType.Warning);
                    }
                }

                gripByGeometry.DeleteDirectShapes();

                byElem = true;
            }

            bool byHost = false;
            if (ElemsByHost.Count > 0)
            {
                byHost = new GripByHost().ExecuteByHostFamily(ElemsByHost, SectionParamName, LevelParamName, pb, gripByGeometry.PbCounter);
            }

            bool byElemsInsulation = false;
            if (ElemsInsulation.Count > 0)
            {
                byElemsInsulation = new GripByHost().ExecuteByElementInsulation(Doc, ElemsInsulation, SectionParamName, LevelParamName, pb, gripByGeometry.PbCounter);
            }

            return byElem && byHost && byElemsInsulation;
        }
    }
}
