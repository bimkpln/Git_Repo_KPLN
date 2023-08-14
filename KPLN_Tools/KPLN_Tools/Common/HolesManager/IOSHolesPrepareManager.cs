using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools.Common.HolesManager
{
    /// <summary>
    /// Класс для обработки спец. класса IOSHoleDTO
    /// </summary>
    internal class IOSHolesPrepareManager
    {
        private readonly IEnumerable<FamilyInstance> _holesElems;
        private readonly IEnumerable<RevitLinkInstance> _linkedModels;

        public IOSHolesPrepareManager(IEnumerable<FamilyInstance> holesElems, IEnumerable<RevitLinkInstance> linkedModels)
        {
            _holesElems = holesElems;
            _linkedModels = linkedModels;
        }

        /// <summary>
        /// Коллекция ошибок, при генерации IOSHoleDTO
        /// </summary>
        public List<FamilyInstance> ErrorFamInstColl { get; private set; } = new List<FamilyInstance>();

        /// <summary>
        /// Подготовка спец. семейст для анализа
        /// </summary>
        /// <param name="holesElems"></param>
        public List<IOSHoleDTO> PrepareHolesDTO()
        {
            List<IOSHoleDTO> result = new List<IOSHoleDTO>();
            foreach (FamilyInstance fi in _holesElems)
            {
                BoundingBoxXYZ fiBBox = PrepareHoleBBox(fi);
                IOSHoleDTO holeDTO = PrepareHoleDTOData(fi, fiBBox);
                if (holeDTO != null)
                    result.Add(holeDTO);
            }

            return result;
        }

        /// <summary>
        /// Подготовка геометрии элемента отверстия
        /// </summary>
        private BoundingBoxXYZ PrepareHoleBBox(FamilyInstance famInst)
        {
            BoundingBoxXYZ result = null;
            GeometryElement geomElem = famInst
                    .get_Geometry(new Options()
                    {
                        DetailLevel = ViewDetailLevel.Fine,
                    });

            foreach (GeometryInstance inst in geomElem)
            {
                //Transform transform = inst.Transform;
                GeometryElement instGeomElem = inst.GetInstanceGeometry();
                foreach (GeometryObject obj in instGeomElem)
                {
                    Solid solid = obj as Solid;
                    if (solid != null && solid.Volume != 0)
                    {
                        BoundingBoxXYZ bbox = solid.GetBoundingBox();
                        //bbox.Transform = transform;
                        Transform transform = bbox.Transform;
                        result = new BoundingBoxXYZ()
                        {
                            Max = transform.OfPoint(bbox.Max),
                            Min = transform.OfPoint(bbox.Min),
                        };

                        return result;
                    }
                }
            }

            throw new Exception($"Не удалось получить геометрию у элемента с id: {famInst.Id}");
        }

        /// <summary>
        /// Подготовка параметров для класса HolesDTO
        /// </summary>
        private IOSHoleDTO PrepareHoleDTOData(FamilyInstance fi, BoundingBoxXYZ fiBBox)
        {
            string fiName = fi.Symbol.FamilyName;

            // Считаю отметки и привязываю HoleDTO
            IOSHoleDTO holesDTO;
            double upMinDist = double.MaxValue;
            double downMinDist = double.MaxValue;
            Floor upFloor = null;
            Floor downFloor = null;
            double downBindElev = 0;
            double rlvDist = 0;

            foreach (RevitLinkInstance linkedModel in _linkedModels)
            {
                Document linkDoc = linkedModel.GetLinkDocument();
                if (linkDoc == null)
                    continue;

                Transform trans = linkedModel.GetTotalTransform();
                BoundingBoxIntersectsFilter filter = CreateFilter(fiBBox, trans);

                // Перевод координат отверстия на координаты связи
                BoundingBoxXYZ inversedFiBBox = new BoundingBoxXYZ()
                {
                    Min = trans.Inverse.OfPoint(fiBBox.Min),
                    Max = trans.Inverse.OfPoint(fiBBox.Max),
                };

                // Поиск перекрытий, которые пересекаются с расширенным отверстием
                IEnumerable<Floor> floorColl = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(Floor))
                    .WherePasses(filter)
                    .Where(fl => fl.Name.StartsWith("00_"))
                    .Cast<Floor>();

                if (floorColl.Any())
                {
                    // Обработка пересекающихся перекрытий
                    foreach (Floor floor in floorColl)
                    {
                        XYZ upPrjPoint = floor.GetVerticalProjectionPoint(inversedFiBBox.Max, FloorFace.Bottom);
                        if (upPrjPoint == null) continue;
                        double upDistance = inversedFiBBox.Max.DistanceTo(upPrjPoint);
                        if (Math.Round(inversedFiBBox.Max.Z, 1) <= Math.Round(upPrjPoint.Z, 1) && upDistance <= upMinDist)
                        {
                            // Проверяю точку максимума и минимума по X, Y. Причина - отверстия в наружных стенах
                            if (CheckFloorFaces(floor, upPrjPoint) || CheckFloorFaces(floor, new XYZ(inversedFiBBox.Min.X, inversedFiBBox.Min.Y, upPrjPoint.Z)))
                            {
                                upFloor = floor;
                                upMinDist = upDistance;
                            }
                        }

                        XYZ downPrjPoint = floor.GetVerticalProjectionPoint(inversedFiBBox.Min, FloorFace.Top);
                        if (downPrjPoint == null) continue;
                        double downDistance = inversedFiBBox.Min.DistanceTo(downPrjPoint);
                        if (Math.Round(inversedFiBBox.Min.Z, 1) >= Math.Round(downPrjPoint.Z, 1) && downDistance <= downMinDist)
                        {
                            // Проверяю точку максимума и минимума по X, Y. Причина - отверстия в наружных стенах
                            if (CheckFloorFaces(floor, downPrjPoint) || CheckFloorFaces(floor, new XYZ(inversedFiBBox.Max.X, inversedFiBBox.Max.Y, downPrjPoint.Z)))
                            {
                                downFloor = floor;

                                Level downFloorLevel = (Level)linkDoc.GetElement(downFloor.LevelId);
                                downBindElev = downFloorLevel.Elevation;
                                double floorAboveLevHeight = downFloor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsDouble();

                                if ((downDistance + floorAboveLevHeight) > 0)
                                    downMinDist = downDistance + floorAboveLevHeight;
                                else
                                {
                                    downMinDist = downDistance;
                                    rlvDist = downMinDist;
                                }

                                if (downBindElev < inversedFiBBox.Min.Z)
                                    rlvDist = downMinDist;
                                else
                                    rlvDist = Math.Round(inversedFiBBox.Min.Z - downBindElev, 2);
                            }
                        }
                    }
                }
            }

            if (downFloor is null)
            {
                ErrorFamInstColl.Add(fi);
                return null;
            }

            // Для отверстий на кровле
            if (upMinDist == double.MaxValue)
                upMinDist = 0.0;

            // Генерация HoleDTO
            if (fiName.ToLower().Contains("прямоуг") || fiName.ToLower().Contains("tsw"))
            {
                holesDTO = new IOSHoleDTO()
                {
                    CurrentHole = fi,
                    UpFloorBinding = upFloor,
                    UpFloorDistance = upMinDist,
                    DownFloorBinding = downFloor,
                    DownFloorDistance = downMinDist,
                    DownBindingElevation = downBindElev,
                    BindingPrefixString = "Низ на отм.",
                    RlvElevation = rlvDist,
                    AbsElevation = fiBBox.Min.Z,
                };
            }
            else if (fiName.ToLower().Contains("кругл") || fiName.ToLower().Contains("trw"))
            {
                holesDTO = new IOSHoleDTO()
                {
                    CurrentHole = fi,
                    UpFloorBinding = upFloor,
                    UpFloorDistance = upMinDist,
                    DownFloorBinding = downFloor,
                    DownFloorDistance = downMinDist,
                    DownBindingElevation = downBindElev,
                    BindingPrefixString = "Центр на отм.",
                    RlvElevation = rlvDist + Math.Abs((Math.Abs(fiBBox.Max.Z) - Math.Abs(fiBBox.Min.Z)) / 2),
                    AbsElevation = ((fiBBox.Max.Z + fiBBox.Min.Z)) / 2,
                };
            }
            else
                throw new Exception($"Не удалось разделить семейство {fiName} на прямоугольное или круглое");

            return holesDTO;
        }

        /// <summary>
        /// Проверка перекрытия на факт пересечения с точкой проецирования
        /// </summary>
        private bool CheckFloorFaces(Floor floor, XYZ prjPoint)
        {
            GeometryElement flGeomElem = floor
                .get_Geometry(new Options()
                {
                    DetailLevel = ViewDetailLevel.Fine,
                });

            foreach (GeometryObject obj in flGeomElem)
            {
                Solid flSolid = obj as Solid;
                if (flSolid != null && flSolid.Volume != 0)
                {
                    FaceArray flFaceArray = flSolid.Faces;
                    foreach (Face flFace in flFaceArray)
                    {
                        IntersectionResult intRes = flFace.Project(prjPoint);
                        if (intRes != null && intRes.Distance == 0)
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
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
                transform.Inverse.OfPoint(new XYZ(sminX - 1, sminY - 1, bbox.Min.Z - 50)),
                transform.Inverse.OfPoint(new XYZ(smaxX + 1, smaxY + 1, bbox.Max.Z + 50)));

            return new BoundingBoxIntersectsFilter(outline);
        }
    }
}
