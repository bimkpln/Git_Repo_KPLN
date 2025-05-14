using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_OpeningHoleManager.Core.MainEntity;
using KPLN_OpeningHoleManager.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_OpeningHoleManager.Core
{
    internal sealed class AROpeningHoleEntity : OpeningHoleEntity
    {
        private bool _ar_OHE_IsHostElementKR = false;

        internal AROpeningHoleEntity(Element elem)
        {
            OHE_Element = elem;
        }

        internal AROpeningHoleEntity(OpenigHoleShape shape, string iosSubDepCode, Element hostElem, XYZ point)
        {
            OHE_Shape = shape;
            AR_OHE_IOSDubDepCode = iosSubDepCode;
            AR_OHE_HostElement = hostElem;
            OHE_Point = point;
        }

        internal AROpeningHoleEntity(OpenigHoleShape shape, string iosSubDepCode, Element hostElem, XYZ point, Element elem) : this(shape, iosSubDepCode, hostElem, point)
        {
            OHE_Element = elem;
        }

        /// <summary>
        /// Элемент-основа для отверстия
        /// </summary>
        internal Element AR_OHE_HostElement { get; private set; }

        /// <summary>
        /// Элемент-основа для отверстия является элементом КР
        /// </summary>
        internal bool AR_OHE_IsHostElementKR 
        {
            get
            {
                if (AR_OHE_HostElement != null)
                {
                    if (AR_OHE_HostElement is Wall wall)
                    {
                        if (wall.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().StartsWith("00_"))
                            _ar_OHE_IsHostElementKR = true;
                    }
                }

                return _ar_OHE_IsHostElementKR;
            }
        }

        /// <summary>
        /// Код раздела, от которого падает задание
        /// </summary>
        internal string AR_OHE_IOSDubDepCode { get; private set; }

        /// <summary>
        /// Создать объединённое отверстие по выбранной коллекции
        /// </summary>
        internal static AROpeningHoleEntity CreateUnionOpeningHole(Document doc, AROpeningHoleEntity[] arOHEColl)
        {
            // Подбираю результирующий тип для семейства
            string resultSubDep = string.Empty;
            IEnumerable<string> arOHESubDeps = arOHEColl.Select(ohe => ohe.AR_OHE_IOSDubDepCode);
            if (arOHESubDeps.Distinct().Count() > 1 || arOHESubDeps.All(subDep => subDep.Equals("Несколько категорий")))
                resultSubDep = "Несколько категорий";
            else
                resultSubDep = arOHESubDeps.FirstOrDefault();


            // Анализирую на наличие нескольких основ у выборки отверстий
            IEnumerable<int> arOHEHostId = arOHEColl.Select(ohe => ohe.AR_OHE_HostElement.Id.IntegerValue);
            if (arOHEHostId.Distinct().Count() > 1)
            {
                HtmlOutput.Print(
                    $"Выбранные отверстия относятся к разным основаниям. Можно объединять отверстия ТОЛЬКО в рамках одной стены. Проанализируй корректность, и выполни объединение вручную",
                    MessageType.Error);
                return null;
            }

            // Получаю единственную основу
            Element hostElem = doc.GetElement(new ElementId(arOHEHostId.FirstOrDefault()));

            // Анализирую сущности и нахожу результирующий размер
            Solid unionSolid = null;
            foreach (AROpeningHoleEntity arOHE in arOHEColl)
            {
                try
                {
                    if (unionSolid == null)
                        unionSolid = arOHE.OHE_Solid;
                    else
                    {
                        Solid tempUnionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(unionSolid, arOHE.OHE_Solid, BooleanOperationsType.Union);
                        if (tempUnionSolid != null && tempUnionSolid.Volume > 0)
                            unionSolid = tempUnionSolid;

                    }
                }
                catch (Exception ex)
                {
                    // Актуально для семейств, у которых нет тела (т.е. просто отверстия, например для СЕТ)
                    if (ex.Message.Contains("проблемы с получением Solid. Отправь разработчику"))
                    {

                    }
                    else
                        throw;
                }
            }

            // Анализирую вектор основы
            XYZ hostDir = GeometryWorker.GetHostDirection(hostElem);

            // Получаю ширину и высоту
            double[] widthAndHeight = GeometryWorker.GetSolidWidhtAndHeight_ByDirection(unionSolid, hostDir);
            double resultWidth = widthAndHeight[0];
            double resultHeight = widthAndHeight[1];


            // Создаю точку вставки
            XYZ unionSolidCentroid = unionSolid.ComputeCentroid();
            BoundingBoxXYZ unionSolidBBox = unionSolid.GetBoundingBox();
            Transform bboxTrans = unionSolidBBox.Transform;

            double locPointX = (bboxTrans.OfPoint(unionSolidBBox.Min).X + bboxTrans.OfPoint(unionSolidBBox.Max).X) / 2;
            double locPointY = (bboxTrans.OfPoint(unionSolidBBox.Min).Y + bboxTrans.OfPoint(unionSolidBBox.Max).Y) / 2;
            double locPointZ = bboxTrans.OfPoint(unionSolidBBox.Min).Z;
            XYZ locPoint = new XYZ(locPointX, locPointY, locPointZ);


            // Создаю сущность для заполнения
            AROpeningHoleEntity result = new AROpeningHoleEntity(
                OpenigHoleShape.Rectangle,
                resultSubDep,
                hostElem,
                locPoint);

            result.SetFamilyPathAndName(doc);
            result.SetGeomParams();
            result.SetGeomParamsRoundData(resultHeight, resultWidth, 0);

            return result;
        }

        /// <summary>
        /// Очистить коллекцию от ложных инстансов по пересечению
        /// </summary>
        internal static AROpeningHoleEntity[] GetEntitesToDel_ByIntescect(Document doc, AROpeningHoleEntity[] arOHEColl)
        {
            List<AROpeningHoleEntity> elemToClearColl = new List<AROpeningHoleEntity>();

            foreach (AROpeningHoleEntity unionOHEEnt in arOHEColl)
            {
                AROpeningHoleEntity[] hostOHEColl = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .Where(el => el is FamilyInstance fi
                        && (fi.Symbol.FamilyName.StartsWith("199_Отвер") || fi.Symbol.FamilyName.StartsWith("ASML_АР_Отверстие"))
                        && fi.Host.Id.IntegerValue == unionOHEEnt.AR_OHE_HostElement.Id.IntegerValue)
                    .Select(el => new AROpeningHoleEntity(el))
                    .ToArray();

                foreach (AROpeningHoleEntity checkHostOHE in hostOHEColl)
                {
                    Solid hostOHESolid = GeometryWorker.GetRevitElemSolid(checkHostOHE.OHE_Element);
                    if (hostOHESolid == null)
                        continue;


                    if ((unionOHEEnt.OHE_Element != null && checkHostOHE.OHE_Element != null)
                        && (unionOHEEnt.OHE_Element.Id.IntegerValue == checkHostOHE.OHE_Element.Id.IntegerValue))
                        continue;

                    //var centr1 = unionOHEEnt.OHE_Solid.ComputeCentroid();
                    //var bbox1 = unionOHEEnt.OHE_Solid.GetBoundingBox();
                    //var trans1 = bbox1.Transform;
                    //var newbbox1 = new BoundingBoxXYZ() { Min = trans1.OfPoint(bbox1.Min), Max = trans1.OfPoint(bbox1.Max)};

                    //var idd2 = checkHostOHE.OHE_Element.Id;
                    //var centr2 = checkHostOHE.OHE_Solid.ComputeCentroid();
                    //var bbox2 = checkHostOHE.OHE_Solid.GetBoundingBox();
                    //var trans2 = bbox2.Transform;
                    //var newbbox2 = new BoundingBoxXYZ() { Min = trans2.OfPoint(bbox2.Min), Max = trans2.OfPoint(bbox2.Max) };

                    Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(hostOHESolid, unionOHEEnt.OHE_Solid, BooleanOperationsType.Intersect);
                    if (intersectSolid != null
                        && intersectSolid.Volume > 0 
                        && !elemToClearColl.Any(ohe => ohe.OHE_Element.Id.IntegerValue == checkHostOHE.OHE_Element.Id.IntegerValue))
                        elemToClearColl.Add(checkHostOHE);
                }
            }

            return elemToClearColl.ToArray();
        }

        /// <summary>
        /// Установить путь к Revit семействам
        /// </summary>
        internal OpeningHoleEntity SetFamilyPathAndName(Document doc)
        {
            if (doc.Title.Contains("СЕТ_1"))
            {
                OHE_FamilyPath_Rectangle = @"X:\BIM\3_Семейства\8_Библиотека семейств Самолета\01_АР\Отверстия, шахты, решетки\ASML_АР_Отверстие прямоугольное.rfa";
                OHE_FamilyPath_Circle = @"X:\BIM\3_Семейства\8_Библиотека семейств Самолета\01_АР\Отверстия, шахты, решетки\ASML_АР_Отверстие круглое.rfa";
            }
            else
            {
                OHE_FamilyPath_Rectangle = @"X:\BIM\3_Семейства\1_АР\199_Отверстия, ниши и шахты\199_Отверстие прямоугольное_(Об_Стена).rfa";
                OHE_FamilyPath_Circle = @"X:\BIM\3_Семейства\1_АР\199_Отверстия, ниши и шахты\199_Отверстие круглое_(Об_Стена).rfa";
            }

            OHE_FamilyName_Rectangle = OHE_FamilyPath_Rectangle
                .Split('\\')
                .FirstOrDefault(part => part.Contains(".rfa"))
                .TrimEnd(".rfa".ToCharArray());

            OHE_FamilyName_Circle = OHE_FamilyPath_Circle
                .Split('\\')
                .FirstOrDefault(part => part.Contains(".rfa"))
                .TrimEnd(".rfa".ToCharArray());

            return this;
        }

        /// <summary>
        /// Установить основные геометрические параметры (ширина, высота, диамтер)
        /// </summary>
        internal AROpeningHoleEntity SetGeomParams()
        {
            if (OHE_Shape == OpenigHoleShape.Rectangle)
            {
                if (OHE_FamilyName_Rectangle.Contains("199_Отверстие прямоугольное"))
                {
                    OHE_ParamNameHeight = "Высота";
                    OHE_ParamNameWidth = "Ширина";
                    OHE_ParamNameExpander = "Расширение границ";
                }
                else if (OHE_FamilyName_Rectangle.Contains("ASML_АР_Отверстие прямоугольное"))
                {
                    OHE_ParamNameHeight = "АР_Высота Проема";
                    OHE_ParamNameWidth = "АР_Ширна Проема";
                }
            }
            else
            {
                if (OHE_FamilyName_Circle.Contains("199_Отверстие круглое"))
                {
                    OHE_ParamNameRadius = "КП_Р_Высота";
                    OHE_ParamNameExpander = "Расширение границ";
                }
                else if (OHE_FamilyName_Circle.Contains("ASML_АР_Отверстие круглое"))
                {
                    OHE_ParamNameRadius = "АР_Высота Проема";
                }
            }

            return this;
        }

        /// <summary>
        /// Установить основные геометрические параметры (ширина, высота, диамтер) И ОКРГУЛИТЬ с шагом 50 мм
        /// </summary>
        internal AROpeningHoleEntity SetGeomParamsRoundData(double height, double width, double radius, double expandValue = 0)
        {
#if Debug2020 || Revit2020
            double roundHeight = RoundGeomParam(height) + RoundGeomParam(UnitUtils.ConvertToInternalUnits(expandValue, DisplayUnitType.DUT_MILLIMETERS));
            double roundWidh = RoundGeomParam(width) + RoundGeomParam(UnitUtils.ConvertToInternalUnits(expandValue, DisplayUnitType.DUT_MILLIMETERS));
            double roundRadius = RoundGeomParam(radius) + RoundGeomParam(UnitUtils.ConvertToInternalUnits(expandValue, DisplayUnitType.DUT_MILLIMETERS));
#else
            double roundHeight = RoundGeomParam(height) + RoundGeomParam(UnitUtils.ConvertToInternalUnits(expandValue, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1")));
            double roundWidh = RoundGeomParam(width) + RoundGeomParam(UnitUtils.ConvertToInternalUnits(expandValue, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1")));
            double roundRadius = RoundGeomParam(radius) + RoundGeomParam(UnitUtils.ConvertToInternalUnits(expandValue, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1")));
#endif

            if (OHE_Shape == OpenigHoleShape.Rectangle)
            {
                if (OHE_FamilyName_Rectangle.Contains("199_Отверстие прямоугольное"))
                {
                    OHE_Height = roundHeight;
                    OHE_Width = roundWidh;
                }
                else if (OHE_FamilyName_Rectangle.Contains("ASML_АР_Отверстие прямоугольное"))
                {
                    OHE_Height = roundHeight;
                    OHE_Width = roundWidh;
                }
            }
            else
            {
                if (OHE_FamilyName_Circle.Contains("199_Отверстие круглое"))
                    OHE_Radius = roundRadius;
                else if (OHE_FamilyName_Circle.Contains("ASML_АР_Отверстие круглое"))
                    OHE_Radius = roundRadius;
            }

            return this;
        }

        /// <summary>
        /// Создать СОЛИД используя параметры AROpeningHoleEntity
        /// </summary>
        internal AROpeningHoleEntity CreateSolid_ByParams()
        {
            OHE_Solid = GeometryWorker.CreateSolid_ZDir(GeometryWorker.GetHostDirection(AR_OHE_HostElement), OHE_Point, OHE_Height, OHE_Width, OHE_Radius);

            return this;
        }

        /// <summary>
        /// Уточнить значения точки в зависимости от формы
        /// </summary>
        /// <returns></returns>
        internal AROpeningHoleEntity UpdatePointData_ByShape()
        {
            XYZ iosTransPnt = this.OHE_Point;

            if (this.OHE_Shape == OpenigHoleShape.Rectangle)
                OHE_Point = new XYZ(iosTransPnt.X, iosTransPnt.Y, iosTransPnt.Z - this.OHE_Height / 2);
            else
                OHE_Point = new XYZ(iosTransPnt.X, iosTransPnt.Y, iosTransPnt.Z - this.OHE_Radius / 2);

            return this;
        }

        /// <summary>
        /// Разместить экземпляр семейства по указанным координатам и заполнить параметры в модели
        /// </summary>
        internal void CreateIntersectFamInstAndSetRevitParamsData(Document doc, Element host)
        {
            // Определяю ключевую часть имени для поиска нужного типа семейства
            string famType = string.Empty;
            if (AR_OHE_IOSDubDepCode.Equals("ОВиК"))
                famType = "ОВ";
            else if (AR_OHE_IOSDubDepCode.Equals("ИТП"))
                famType = "ОВ";
            else if (AR_OHE_IOSDubDepCode.Equals("ВК"))
                famType = "ВК";
            else if (AR_OHE_IOSDubDepCode.Equals("ПТ"))
                famType = "ВК";
            else if (AR_OHE_IOSDubDepCode.Equals("ЭОМ"))
                famType = "ЭОМ";
            else if (AR_OHE_IOSDubDepCode.Equals("СС"))
                famType = "СС";
            else if (AR_OHE_IOSDubDepCode.Equals("АВ"))
                famType = "СС";
            else
                famType = "Несколько категорий";

            FamilySymbol openingFamSymb;
            if (OHE_Shape == OpenigHoleShape.Rectangle)
                openingFamSymb = GetIntersectFamilySymbol(doc, OHE_FamilyPath_Rectangle, OHE_FamilyName_Rectangle, famType);
            else
                openingFamSymb = GetIntersectFamilySymbol(doc, OHE_FamilyPath_Circle, OHE_FamilyName_Circle, famType);

            if (host.LevelId == null)
                throw new Exception($"У основы с id: {host.Id} проблемы с привязкой к уровню. Отправь разработчику.");

            Level hostLevel = doc.GetElement(host.LevelId) as Level;

            // Создание новых экземпляров
            FamilyInstance instance = doc
                .Create
                .NewFamilyInstance(OHE_Point, openingFamSymb, host, hostLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

            // Присваиваю параметр эл-та модели инстансу класса (далее используется)
            OHE_Element = instance;
            doc.Regenerate();

            // Указать уровень - для семейств на основе указывать НЕ нужно
            //instance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(OHE_Point.Z - hostLevel.Elevation);

            // Заполнить параметры
            if (OHE_Shape == OpenigHoleShape.Rectangle)
            {
                instance.LookupParameter(OHE_ParamNameHeight).Set(OHE_Height);
                instance.LookupParameter(OHE_ParamNameWidth).Set(OHE_Width);

                Parameter expandParam = instance.LookupParameter(OHE_ParamNameExpander);
                if (expandParam != null && !expandParam.IsReadOnly)
                    instance.LookupParameter(OHE_ParamNameExpander).Set(0);
            }
            else
            {
                instance.LookupParameter(OHE_ParamNameRadius).Set(OHE_Radius);

                Parameter expandParam = instance.LookupParameter(OHE_ParamNameExpander);
                if (expandParam != null && !expandParam.IsReadOnly)
                    instance.LookupParameter(OHE_ParamNameExpander).Set(0);
            }


            doc.Regenerate();
            
            // Присваиваю солид ПОСЛЕ установки параметров
            OHE_Solid = GeometryWorker.GetRevitElemSolid(instance);
        }

        private double RoundGeomParam(double geomParam)
        {
            double round_mm;
#if Debug2020 || Revit2020
            double mm = UnitUtils.ConvertFromInternalUnits(geomParam, DisplayUnitType.DUT_MILLIMETERS);
#else
            double mm = UnitUtils.ConvertFromInternalUnits(geomParam, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1"));
#endif
            if (mm % 50 < 0.1)
                round_mm = Math.Round(mm);
            else
                round_mm = Math.Ceiling(mm / 50) * 50;

#if Debug2020 || Revit2020
            return UnitUtils.ConvertToInternalUnits(round_mm, DisplayUnitType.DUT_MILLIMETERS);
#else
            return UnitUtils.ConvertToInternalUnits(round_mm, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1"));
#endif
        }
    }
}
