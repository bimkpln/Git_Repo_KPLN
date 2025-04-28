using Autodesk.Revit.DB;
using KPLN_OpeningHoleManager.Core.MainEntity;
using System;
using System.Linq;
using System.Windows.Media.Media3D;

namespace KPLN_OpeningHoleManager.Core
{
    internal sealed class AROpeningHoleEntity : OpeningHoleEntity
    {
        public AROpeningHoleEntity(OpenigHoleShape shape, string iosSubDepCode, Element hostElem, XYZ point)
        {
            OHE_Shape = shape;
            AR_OHE_IOSDubDepCode = iosSubDepCode;
            AR_OHE_HostElement = hostElem;
            OHE_Point = point;
        }

        public AROpeningHoleEntity(OpenigHoleShape shape, string iosSubDepCode, Element hostElem, XYZ point, Element elem) : this(shape, iosSubDepCode, hostElem, point)
        {
            OHE_Element = elem;
        }

        /// <summary>
        /// Элемент-основа для отверстия
        /// </summary>
        public Element AR_OHE_HostElement { get; private set; }

        /// <summary>
        /// Код раздела, от которого падает задание
        /// </summary>
        public string AR_OHE_IOSDubDepCode { get; private set; }

        /// <summary>
        /// Установить путь к Revit семействам
        /// </summary>
        public OpeningHoleEntity SetFamilyPathAndName(Document doc)
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
        public AROpeningHoleEntity SetGeomParams()
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
        public AROpeningHoleEntity SetGeomParamsRoundData(double height, double width, double radius)
        {
            double roundHeight = RoundGeomParam(height);
            double roundWidh = RoundGeomParam(width);
            double roundRadius = RoundGeomParam(radius);


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
        /// Уточнить значения точки по IOSTask
        /// </summary>
        /// <returns></returns>
        public AROpeningHoleEntity UpdatePointData(IOSOpeningHoleTaskEntity iosTask)
        {
            XYZ iosTransPnt = iosTask.OHE_LinkTransform.OfPoint(iosTask.OHE_Point);

            if (iosTask.OHE_Shape == OpenigHoleShape.Rectangle)
                OHE_Point = new XYZ(iosTransPnt.X, iosTransPnt.Y, iosTransPnt.Z - iosTask.OHE_Height / 2);
            else
                OHE_Point = new XYZ(iosTransPnt.X, iosTransPnt.Y, iosTransPnt.Z - iosTask.OHE_Radius / 2);

            return this;
        }

        /// <summary>
        /// Разместить экземпляр семейства по указанным координатам и заполнить параметры в модели
        /// </summary>
        public void CreateIntersectFamInstAndSetRevitParamsData(Document doc, Element host)
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

            doc.Regenerate();

            // Указать уровень
            instance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(OHE_Point.Z - hostLevel.Elevation);

            // Заполнить параметры
            if (OHE_Shape == OpenigHoleShape.Rectangle)
            {
                instance.LookupParameter(OHE_ParamNameHeight).Set(OHE_Height);
                instance.LookupParameter(OHE_ParamNameWidth).Set(OHE_Width);

                Parameter expandParam = instance.LookupParameter(OHE_ParamNameExpander);
                if (expandParam != null)
                    instance.LookupParameter(OHE_ParamNameExpander).Set(0);
            }
            else
            {
                instance.LookupParameter(OHE_ParamNameRadius).Set(OHE_Radius);
                Parameter expandParam = instance.LookupParameter(OHE_ParamNameExpander);
                if (expandParam != null)
                    instance.LookupParameter(OHE_ParamNameExpander).Set(0);
            }

            doc.Regenerate();
        }

        private double RoundGeomParam(double geomParam)
        {
            double round_mm;
            double mm = UnitUtils.ConvertFromInternalUnits(geomParam, DisplayUnitType.DUT_MILLIMETERS);
            if (mm % 50 < 0.1)
                round_mm = Math.Round(mm);
            else
                round_mm = Math.Ceiling(mm / 50) * 50;

            return UnitUtils.ConvertToInternalUnits(round_mm, DisplayUnitType.DUT_MILLIMETERS);
        }
    }
}
