using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using KPLN_Parameters_Ribbon.Common.Tools;
using KPLN_Parameters_Ribbon.Forms;
using System;
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

        public override bool Prepare()
        {
            List<BuiltInCategory> userCat = null;
            List<BuiltInCategory> revitCat = null;
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
                    BuiltInCategory.OST_GenericModel
                };
                
                // Категории системных семейств, используемые в проектах ИОС
                revitCat = new List<BuiltInCategory>()
                {
                    BuiltInCategory.OST_CableTray,
                    BuiltInCategory.OST_Conduit,
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
                    BuiltInCategory.OST_DuctInsulations,
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_FlexPipeCurves,
                    BuiltInCategory.OST_PipeInsulations,
                };
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
                            .Where(x => !x.Name.StartsWith("500_"))
                            .Where(x => !x.Name.StartsWith("501_"))
                            .Where(x => !x.Name.StartsWith("502_"))
                            .Where(x => !x.Name.StartsWith("503_")));
                        break;
                    
                    case BuiltInCategory.OST_GenericModel:
                        _dirtyElems.AddRange(new FilteredElementCollector(Doc)
                            .OfClass(typeof(FamilyInstance))
                            .OfCategory(bic)
                            .Cast<FamilyInstance>()
                            .Where(x => !x.Name.StartsWith("500_"))
                            .Where(x => !x.Name.StartsWith("501_"))
                            .Where(x => !x.Name.StartsWith("502_"))
                            .Where(x => !x.Name.StartsWith("503_")));
                        break;
                    
                    default:
                        _dirtyElems.AddRange(new FilteredElementCollector(Doc)
                            .OfClass(typeof(FamilyInstance))
                            .OfCategory(bic)
                            .Cast<FamilyInstance>());
                        break;
                }
            }

            foreach (BuiltInCategory bic in revitCat)
            {
                ElemsOnLevel.AddRange(new FilteredElementCollector(Doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType());
            }

            ElemsOnLevel.AddRange(_dirtyElems
                .Where(x => x.SuperComponent == null));

            ElemsByHost.AddRange(_dirtyElems
                .Where(x => x.SuperComponent != null));

            AllElementsCount = ElemsOnLevel.Count + ElemsByHost.Count;

            if (AllElementsCount > 0)
            {
                return true;
            }
            else
            {
                throw new Exception("KPLN: Ошибка при взятии элементов из проекта. Таких категорий нет, или имя проекта не соответсвует ВЕР!");
            }
        }


        public override bool ExecuteGripParams(Progress_Single pb)
        {
            MainTool mainTool = new MainTool(Doc, LevelParamName, SectionParamName);

            IEnumerable<MyLevel> myLevelColl = mainTool.LevelPrepare();

            Dictionary<string, HashSet<Grid>> gridsDict = mainTool.GridPrepare("КП_О_Секция");

            IEnumerable<MySolid> mySolids = mainTool.SolidsCollectionPrepare(gridsDict, myLevelColl, LevelNumberIndex, SplitLevelChar);

            bool byElem = false;
            if (ElemsOnLevel.Count > 0)
            {

                IEnumerable<Element> notIntersectedElems = mainTool.IntersectWithSolidExcecute(ElemsOnLevel, mySolids, pb);

                IEnumerable<Element> notNearestSolidElems = mainTool.FindNearestSolid(notIntersectedElems, mySolids, pb);

                IEnumerable<Element> notRevalueElems = mainTool.ReValueDuplicates(mySolids, pb);

                if (mainTool.DuplicatesWriteParamElems.Keys.Count() > 0)
                {
                    foreach (Element element in mainTool.DuplicatesWriteParamElems.Keys)
                    {
                        Print($"Проверь вручную элемент с id: {element.Id}", KPLN_Loader.Preferences.MessageType.Warning);
                    }
                }

                mainTool.DeleteDirectShapes();
                
                byElem = true;
            }

            return byElem;
        }

        public override bool ExecuteLevelParams(Progress_Single pb)
        {
            bool byElem = false;
            bool byHost = false;

            var a = ExecuteGripParams(pb);

            //if (ElemsOnLevel.Count > 0)
            //{
            //    byElem = LevelExcecuter.ExecuteByElement(Doc, ElemsOnLevel, "КП_О_Секция", LevelParamName, pb);
            //}

            //if (ElemsByHost.Count > 0)
            //{
            //    byHost = LevelExcecuter.ExecuteByHost(ElemsByHost, LevelParamName, pb);
            //}

            return byElem && byHost;

            //LevelTool.Levels = new FilteredElementCollector(Doc)
            //        .OfClass(typeof(Level))
            //        .WhereElementIsNotElementType()
            //        .Cast<Level>();

            //FloorNumberOnLevelByElement(pb);

            //FloorNumberOnLevelByHost(pb);

            //return true;
        }

        public override bool ExecuteSectionParams(Progress_Single pb)
        {
            bool byElem = false;
            bool byHost = false;

            if (ElemsOnLevel.Count > 0)
            {
                byElem = SectionExcecuter.ExecuteByElement(Doc, ElemsOnLevel, "КП_О_Секция", SectionParamName, pb);
            }

            if (ElemsByHost.Count > 0)
            {
                byHost = SectionExcecuter.ExecuteByHost(ElemsByHost, SectionParamName, pb);
            }
            
            return byElem && byHost;
        }

        /// <summary>
        /// Заполняю номер этажа для элементов, находящихся НА уровне
        /// </summary>
        protected virtual void FloorNumberOnLevelByElement(Progress_Single pb)
        {
            foreach (Element elem in ElemsOnLevel)
            {
                try
                {
                    XYZ elemPoint = null;
                    Location location = elem.Location;
                    Type locationType = location.GetType();
                    switch (locationType.Name)
                    {
                        case nameof(LocationCurve):
                            LocationCurve locCurve = location as LocationCurve;
                            Curve curve = locCurve.Curve;
                            elemPoint = (curve.GetEndPoint(0) + curve.GetEndPoint(1))/2;
                            break;
                        case nameof(LocationPoint):
                            LocationPoint locationPoint = location as LocationPoint;
                            elemPoint = locationPoint.Point;
                            break;
                        default:
                            throw new Exception("Не удалось получить отметку у элемента с Id:" + elem.Id.IntegerValue.ToString());
                    }

                    // СЛАБОЕ МЕСТО ДЛЯ РАЗНОСЕКЦИОННЫХ ПРОЕКТОВ - МОЖЕТ ВПИСАТЬ НЕ ТОТ УРОВЕНЬ ПРИ БОЛЬШИХ ПЕРЕПАДАХ МЕЖДУ СЕКЦИЯМИ
                    Level currentLevel = GetNearestBelowPointLevel(elemPoint, Doc);
                    string floorNumber = GetFloorNumberByLevel(currentLevel, LevelNumberIndex, SplitLevelChar);
                    
                    if (floorNumber == null) continue;
                    Parameter floor = elem.LookupParameter(LevelParamName);
                    if (floor == null) continue;
                    floor.Set(floorNumber);

                    pb.Increment();
                }
                catch (Exception e)
                {
                    PrintError(e, $"Не удалось обработать элемент: {elem.Id.IntegerValue} / {elem.Name}");
                }
            }
        }

        /// <summary>
        /// Заполняю номер этажа для вложенных элементов
        /// </summary>
        protected virtual void FloorNumberOnLevelByHost(Progress_Single pb)
        {
            foreach (Element elem in ElemsByHost)
            {
                FamilyInstance instance = elem as FamilyInstance;
                Element hostElem = instance.SuperComponent;
                if (hostElem == null)
                {
                    throw new Exception($"Невозможно взять родительское семейство для: {elem.Id.IntegerValue} / {elem.Name}");
                }
                var hostElemParamValue = hostElem.LookupParameter(LevelParamName).AsString();
                elem.LookupParameter(LevelParamName).Set(hostElemParamValue);

                pb.Increment(); ;
            }
        }

        /// <summary>
        /// Поиск уровня, ближайшего к точке и ниже этой точки. Для элементов на нижних уровнях с отрицательной отметкой выдаст нижний уровень
        /// </summary>
        /// <param name="point">Точка в пространстве</param>
        /// <param name="doc">Документ для анализа</param>
        /// <returns>Id уровня</returns>
        public static Level GetNearestBelowPointLevel(XYZ point, Document doc)
        {
            var Levels = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .WhereElementIsNotElementType()
                    .Cast<Level>();


            double pointZ = point.Z;

            BasePoint projectBasePoint = new FilteredElementCollector(doc)
                .OfClass(typeof(BasePoint))
                .WhereElementIsNotElementType()
                .Cast<BasePoint>()
                .Where(i => i.IsShared == false)
                .First();
            double projectPointElevation = projectBasePoint.get_BoundingBox(null).Min.Z;

            Level nearestLevel = null;

            double offset = 10000;
            foreach (Level lev in Levels)
            {
                double levHeigth = lev.Elevation + projectPointElevation;
                double tempElev = pointZ - levHeigth;
                if (tempElev < 0) continue;

                if (tempElev < offset)
                {
                    nearestLevel = lev;
                    offset = tempElev;
                }
            }

            if (nearestLevel == null)
            {
                foreach (Level lev in Levels)
                {
                    if (nearestLevel == null)
                    {
                        nearestLevel = lev;
                        continue;
                    }
                    if (lev.Elevation < nearestLevel.Elevation)
                    {
                        nearestLevel = lev;
                    }
                }
            }

            return nearestLevel;
        }

        /// <summary>
        /// Поиск номера уровня у элемента (над уровнем) по уровню
        /// </summary>
        /// <param name="lev">Уровень</param>
        /// <param name="floorTextPosition">Индекс номера уровня с учетом разделителя</param>
        /// <param name="splitChar">Разделитель</param>
        /// <returns>Номер уровня</returns>
        public static string GetFloorNumberByLevel(Level lev, int floorTextPosition, char splitChar)
        {
            string levname = lev.Name;
            string[] splitname = levname.Split(splitChar);
            if (splitname.Length < 2)
            {
                throw new Exception($"Некорректное имя уровня: {levname}");
            }
            string floorNumber = splitname[floorTextPosition];

            return floorNumber;
        }
    }
}
