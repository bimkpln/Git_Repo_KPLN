using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Tools.Common.HolesManager
{
    internal class ShaftPrepareManager
    {
        /// <summary>
        /// Подготовка спец. семейст для анализа
        /// </summary>
        public static List<ShaftDTO> PrepareShaftDTO(IEnumerable<FamilyInstance> shaftElems, IEnumerable<RevitLinkInstance> linkedModels, double absoluteElevBasePnt)
        {
            List<ShaftDTO> result = new List<ShaftDTO>();
            foreach (FamilyInstance fi in shaftElems)
            {
                string fiName = fi.Symbol.FamilyName;

                // Считаю отметки и привязываю ShaftDTO
                ShaftDTO shaftDTO = null;
                Floor downFloor = null;
                double downBindElev = 0;
                double rlvDist = 0;

                BoundingBoxXYZ fiBBox = fi
                        .get_Geometry(new Options()
                        {
                            DetailLevel = ViewDetailLevel.Fine,
                        })
                        .GetBoundingBox();
                BoundingBoxXYZ expandedBBox = new BoundingBoxXYZ()
                {
                    Min = new XYZ (fiBBox.Max.X, fiBBox.Max.Y, fiBBox.Min.Z - 1),
                    Max = new XYZ (fiBBox.Max.X, fiBBox.Max.Y, fiBBox.Max.Z + 1),
                };

                // Создание расширенного контура, для поиска перекрытий, над/под которыми оно расположено
                BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(new Outline(
                    new XYZ(expandedBBox.Max.X, expandedBBox.Max.Y, expandedBBox.Min.Z),
                    new XYZ(expandedBBox.Max.X, expandedBBox.Max.Y, expandedBBox.Max.Z)));

                foreach (RevitLinkInstance linkedModel in linkedModels)
                {
                    // Поиск перекрытий, которые пересекаются с расширенным отверстием
                    IEnumerable<Floor> floorColl = new FilteredElementCollector(linkedModel.GetLinkDocument())
                        .OfClass(typeof(Floor))
                        .WherePasses(filter)
                        .Where(fl => fl.Name.StartsWith("00_"))
                        .Cast<Floor>();

                    // Обработка пересекающихся перекрытий
                    downBindElev = floorColl.Min(f => ((Level)linkedModel.GetLinkDocument().GetElement(f.LevelId)).Elevation);
                    downFloor = floorColl.Where(f => ((Level)linkedModel.GetLinkDocument().GetElement(f.LevelId)).Elevation == downBindElev).FirstOrDefault();
                    rlvDist = fiBBox.Min.Z - downBindElev;
                }

                if (downFloor is null)
                    throw new Exception($"KPLN: Ошибка - экземпляр с id: {fi.Id} невозможно определить основу (уровень, на котором оно расположено)");
                
                shaftDTO = new ShaftDTO()
                {
                    CurrentHole = fi,
                    DownFloorBinding = downFloor,
                    DownBindingElevation = downBindElev,
                    BindingPrefixString = "Низ на отм.",
                    RlvElevation = rlvDist,
                    AbsElevation = absoluteElevBasePnt - Math.Abs(fiBBox.Min.Z),
                };

                result.Add(shaftDTO);
            }

            return result;
        }
    }
}
