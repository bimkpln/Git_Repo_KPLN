using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_Parameters_Ribbon.Common.Tools;
using KPLN_Parameters_Ribbon.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Loader.Output.Output;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Builder
{
    internal class GripBuilder_AR : AbstrGripBuilder
    {
        public GripBuilder_AR(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName)
        {
        }

        public GripBuilder_AR(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName, char splitLevelChar) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName, splitLevelChar)
        {
        }

        public override void Prepare()
        {
            // Категория "Стены"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(x => !x.Name.StartsWith("00_")));

            // Категория "Перекрытия"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Floor))
                .Cast<Floor>()
                .Where(x => !x.Name.StartsWith("00_")));

            // Категория "Кровля"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(RoofBase))
                .Cast<RoofBase>());

            // Семейства "Окна"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Windows)
                .Cast<FamilyInstance>());

            // Семейства "Двери"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Doors)
                .Cast<FamilyInstance>());

            // Семейства "Панели витража"
            ElemsByHost.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_CurtainWallPanels)
                .Cast<FamilyInstance>());

            // Семейства "Импосты витража"
            ElemsByHost.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_CurtainWallMullions)
                .Cast<FamilyInstance>());

            // Семейства "Лестничные марши"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(Railing))
                .Cast<Railing>());

            // Семейства "Оборудование"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                .Cast<FamilyInstance>()
                .Where(x => !x.Symbol.FamilyName.StartsWith("199_")));

            // Семейства "Обощенные модели"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Cast<FamilyInstance>()
                .Where(x => !x.Symbol.FamilyName.StartsWith("199_")));

            // Семейства "Каркас несущий (перемычки)"
            ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .Cast<FamilyInstance>()
                .Where(x => !x.Symbol.FamilyName.StartsWith("199_")));
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
                byHost = new GripByHost().ExecuteByHostFamily_AR(ElemsByHost, SectionParamName, LevelParamName, pb, gripByGeometry.PbCounter);
            }

            return byElem && byHost;
        }
    }
}
