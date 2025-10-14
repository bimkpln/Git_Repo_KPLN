using Autodesk.Revit.DB;
using KPLN_OpeningHoleManager.Core.MainEntity;
using System.Linq;

namespace KPLN_OpeningHoleManager.Core
{
    /// <summary>
    /// Сущность задания на отверстие от ИОС для АР/КР в модели
    /// </summary>
    internal sealed class IOSOpeningHoleTaskEntity : OpeningHoleEntity
    {
        internal IOSOpeningHoleTaskEntity(Element elem) : base(elem)
        {
        }

        internal IOSOpeningHoleTaskEntity(Element elem, Document doc, XYZ point) : this(elem) 
        {
            OHE_LinkDocument = doc;
            OHE_Point = point;
        }

        /// <summary>
        /// Установить путь к Revit семействам
        /// </summary>
        internal OpeningHoleEntity SetFamilyPathAndName(Document doc)
        {
            if (doc.Title.Contains("СЕТ_1"))
            {
                OHE_FamilyPath_Rectangle = @"X:\BIM\3_Семейства\8_Библиотека семейств Самолета\00_Общие семейства\2_Отверстия\ASML_О_Отверстие_Прямоугольное_В стене.rfa";
                OHE_FamilyPath_Circle = @"X:\BIM\3_Семейства\8_Библиотека семейств Самолета\00_Общие семейства\2_Отверстия\ASML_О_Отверстие_Круглое_В стене.rfa";
            }
            else
            {
                OHE_FamilyPath_Rectangle = @"X:\BIM\3_Семейства\0_Общие семейства\2_ИОС\BIM\501_ЗИ_Отверстие_Прямоугольное_Стена_(Об).rfa";
                OHE_FamilyPath_Circle = @"X:\BIM\3_Семейства\0_Общие семейства\2_ИОС\BIM\501_ЗИ_Отверстие_Круглое_Стена_(Об).rfa";
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
        internal IOSOpeningHoleTaskEntity SetGeomParams()
        {
            if (IEDElem != null)
            {
                if (OHE_Shape == OpenigHoleShape.Rectangular)
                {
                    if (OHE_FamilyName_Rectangle.Contains("501_ЗИ_Отверстие_Прямоугольное"))
                    {
                        OHE_ParamNameHeight = "Высота";
                        OHE_ParamNameWidth = "Ширина";
                        OHE_ParamNameExpander = "Расширение границ";
                    }
                    else if (OHE_FamilyName_Rectangle.Contains("ASML_О_Отверстие_Прямоугольное"))
                    {
                        OHE_ParamNameHeight = "ASML_Размер_Высота";
                        OHE_ParamNameWidth = "ASML_Размер_Ширина";
                        OHE_ParamNameExpander = "Расширение границ";
                    }
                }
                else
                {
                    if (OHE_FamilyName_Circle.Contains("501_ЗИ_Отверстие_Круглое"))
                    {
                        OHE_ParamNameRadius = "КП_Р_Диаметр";
                        OHE_ParamNameExpander = "Расширение границ";
                    }
                    else if (OHE_FamilyName_Circle.Contains("ASML_О_Отверстие_Круглое"))
                    {
                        OHE_ParamNameRadius = "ASML_Размер_Диаметр";
                        OHE_ParamNameExpander = "Расширение границ";
                    }
                }
            }

            return this;
        }

        /// <summary>
        /// Установить данные по основным геометрическим параметрам (ширина, высота, диамтер)
        /// </summary>
        internal IOSOpeningHoleTaskEntity SetGeomParamsData()
        {
            if (IEDElem != null)
            {
                if (OHE_Shape == OpenigHoleShape.Rectangular)
                {
                    if (OHE_FamilyName_Rectangle.Contains("_Прямоугольное_"))
                    {
                        if (OHE_ParamNameHeight == null || OHE_ParamNameWidth == null || OHE_ParamNameExpander == null)
                            throw new System.Exception ("У элемнта заданий нет нужных парамтеров. Обратись к разработчику");

                        OHE_Height = IEDElem.LookupParameter(OHE_ParamNameHeight).AsDouble() + IEDElem.LookupParameter(OHE_ParamNameExpander).AsDouble() * 2;
                        OHE_Width = IEDElem.LookupParameter(OHE_ParamNameWidth).AsDouble() + IEDElem.LookupParameter(OHE_ParamNameExpander).AsDouble() * 2;
                    }
                }
                else
                {
                    if (OHE_FamilyName_Circle.Contains("_Круглое_"))
                    {
                        if (OHE_ParamNameRadius == null || OHE_ParamNameExpander == null)
                            throw new System.Exception("У элемнта заданий нет нужных парамтеров. Обратись к разработчику");
                        
                        OHE_Radius = IEDElem.LookupParameter(OHE_ParamNameRadius).AsDouble();
                    }
                }
            }

            return this;
        }
    }
}
