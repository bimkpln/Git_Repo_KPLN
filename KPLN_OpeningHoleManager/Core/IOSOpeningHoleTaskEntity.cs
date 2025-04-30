using Autodesk.Revit.DB;
using KPLN_OpeningHoleManager.Core.MainEntity;
using System.Linq;

namespace KPLN_OpeningHoleManager.Core
{
    internal sealed class IOSOpeningHoleTaskEntity : OpeningHoleEntity
    {
        public IOSOpeningHoleTaskEntity(Document doc, Element elem, XYZ point)
        {
            OHE_LinkDocument = doc;
            OHE_Element = elem;
            OHE_Point = point;
        }

        /// <summary>
        /// Установить путь к Revit семействам
        /// </summary>
        public OpeningHoleEntity SetFamilyPathAndName(Document doc)
        {
            if (doc.Title.Contains("СЕТ_1"))
            {
                OHE_FamilyPath_Rectangle = @"X:\BIM\3_Семейства\0_Общие семейства\2_ИОС\BIM\501_ЗИ_Отверстие_Прямоугольное_Стена_(Об).rfa";
                OHE_FamilyPath_Circle = @"X:\BIM\3_Семейства\0_Общие семейства\2_ИОС\BIM\501_ЗИ_Отверстие_Круглое_Стена_(Об).rfa";
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
        public IOSOpeningHoleTaskEntity SetGeomParams()
        {
            if (OHE_Element != null)
            {
                if (OHE_Shape == OpenigHoleShape.Rectangle)
                {
                    if (OHE_FamilyName_Rectangle.Contains("501_ЗИ_Отверстие_Прямоугольное"))
                    {
                        OHE_ParamNameHeight = "Высота";
                        OHE_ParamNameWidth = "Ширина";
                    }
                }
                else
                {
                    if (OHE_FamilyName_Circle.Contains("501_ЗИ_Отверстие_Круглое"))
                    {
                        OHE_ParamNameRadius = "КП_Р_Радиус";
                    }
                }
            }

            return this;
        }

        /// <summary>
        /// Установить данные по основным геометрическим параметрам (ширина, высота, диамтер)
        /// </summary>
        public IOSOpeningHoleTaskEntity SetGeomParamsData()
        {
            if (OHE_Element != null)
            {
                if (OHE_Shape == OpenigHoleShape.Rectangle)
                {
                    if (OHE_FamilyName_Rectangle.Contains("501_ЗИ_Отверстие_Прямоугольное"))
                    {
                        if (OHE_ParamNameHeight == null || OHE_ParamNameWidth == null)
                            throw new System.Exception ("У элемнта заданий нет нужных парамтеров. Обратись к разработчику");
                        
                        OHE_Height = OHE_Element.LookupParameter(OHE_ParamNameHeight).AsDouble();
                        OHE_Width = OHE_Element.LookupParameter(OHE_ParamNameWidth).AsDouble();
                    }
                }
                else
                {
                    if (OHE_FamilyName_Circle.Contains("501_ЗИ_Отверстие_Круглое"))
                    {
                        if (OHE_ParamNameRadius == null)
                            throw new System.Exception("У элемнта заданий нет нужных парамтеров. Обратись к разработчику");
                        
                        OHE_Radius = OHE_Element.LookupParameter(OHE_ParamNameRadius).AsDouble() * 2;
                    }
                }
            }

            return this;
        }
    }
}
