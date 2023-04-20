using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Tools.Common.HolesManager
{
    internal class IOSShaftPrepareManager
    {
        private readonly IEnumerable<FamilyInstance> _shaftElems;
        private readonly IEnumerable<RevitLinkInstance> _linkedModels;
        private readonly double _absoluteElevBasePnt;

        public IOSShaftPrepareManager(IEnumerable<FamilyInstance> shaftElems, IEnumerable<RevitLinkInstance> linkedModels, double absoluteElevBasePnt)
        {
            _shaftElems = shaftElems;
            _linkedModels = linkedModels;
            _absoluteElevBasePnt = absoluteElevBasePnt;
        }
        
        /// <summary>
        /// Подготовка спец. семейст для анализа
        /// </summary>
        public List<IOSShaftDTO> PrepareShaftDTO()
        {
            List<IOSShaftDTO> result = new List<IOSShaftDTO>();
            foreach (FamilyInstance fi in _shaftElems)
            {
                string fiName = fi.Symbol.FamilyName;

                // Считаю отметки и привязываю ShaftDTO
                IOSShaftDTO shaftDTO = null;
                Floor downFloor = null;
                double downBindElev = 0;
                double rlvDist = 0;

                BoundingBoxXYZ fiBBox = fi
                        .get_Geometry(new Options()
                        {
                            DetailLevel = ViewDetailLevel.Fine,
                        })
                        .GetBoundingBox();

                foreach (RevitLinkInstance linkedModel in _linkedModels)
                {
                    Document linkDoc = linkedModel.GetLinkDocument();
                    if (linkDoc == null)
                        continue;

                    Transform trans = linkedModel.GetTotalTransform();
                    BoundingBoxIntersectsFilter filter = CreateFilter(fiBBox, trans);

                    // Поиск перекрытий, которые пересекаются с расширенным отверстием
                    IEnumerable<Floor> floorColl = new FilteredElementCollector(linkDoc)
                        .OfClass(typeof(Floor))
                        .WherePasses(filter)
                        .Where(fl => fl.Name.StartsWith("00_"))
                        .Cast<Floor>();
                    List<Floor> floorList = floorColl.ToList();

                    // Обработка пересекающихся перекрытий
                    if (floorList.Any())
                    {
                        downBindElev = floorList.Min(f => ((Level)linkDoc.GetElement(f.LevelId)).Elevation);
                        downFloor = floorList.Where(f => ((Level)linkDoc.GetElement(f.LevelId)).Elevation == downBindElev).FirstOrDefault();
                        rlvDist = fiBBox.Min.Z - downBindElev;
                    }
                }

                if (downFloor is null)
                    throw new Exception($"KPLN: Ошибка - экземпляр с id: {fi.Id} невозможно определить основу (уровень, на котором оно расположено)");
                
                shaftDTO = new IOSShaftDTO()
                {
                    CurrentHole = fi,
                    DownFloorBinding = downFloor,
                    DownBindingElevation = downBindElev,
                    BindingPrefixString = "Низ на отм.",
                    RlvElevation = rlvDist,
                    AbsElevation = _absoluteElevBasePnt - Math.Abs(fiBBox.Min.Z),
                };

                result.Add(shaftDTO);
            }

            return result;
        }

        /// <summary>
        /// Создание фильтра, для поиска перекрытий, над/под которыми оно расположено
        /// </summary>
        private BoundingBoxIntersectsFilter CreateFilter(BoundingBoxXYZ bbox, Transform transform)
        {
            double minX = bbox.Min.X;
            double minY = bbox.Min.Y;

            double maxX = bbox.Max.X;
            double maxY = bbox.Max.Y;

            double sminX = Math.Min(minX, maxX);
            double sminY = Math.Min(minY, maxY);

            double smaxX = Math.Max(minX, maxX);
            double smaxY = Math.Max(minY, maxY);

            Outline outline = new Outline(
                transform.Inverse.OfPoint(new XYZ(sminX, sminY, bbox.Min.Z - 100)),
                transform.Inverse.OfPoint(new XYZ(smaxX, smaxY, bbox.Max.Z + 100)));

            return new BoundingBoxIntersectsFilter(outline);
        }
    }
}
