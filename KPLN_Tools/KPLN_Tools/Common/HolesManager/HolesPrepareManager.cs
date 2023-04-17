using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;

namespace KPLN_Tools.Common.HolesManager
{
    /// <summary>
    /// Класс для обработки спец. класса HoleDTO
    /// </summary>
    internal class HolesPrepareManager
    {
        /// <summary>
        /// Подготовка спец. семейст для анализа
        /// </summary>
        /// <param name="holesElems"></param>
        public static List<HoleDTO> PrepareHolesDTO(IEnumerable<FamilyInstance> holesElems, IEnumerable<RevitLinkInstance> linkedModels, double absoluteElevBasePnt)
        {
            List<HoleDTO> result = new List<HoleDTO>();
            foreach (FamilyInstance fi in holesElems)
            {
                BoundingBoxXYZ fiBBox = PrepareHoleBBox(fi);
                HoleDTO holeDTO = PrepareHoleDTOData(fi, fiBBox, linkedModels, absoluteElevBasePnt);
                result.Add(holeDTO);
            }

            return result;
        }

        /// <summary>
        /// Подготовка геометрии элемента отверстия
        /// </summary>
        private static BoundingBoxXYZ PrepareHoleBBox(FamilyInstance famInst)
        {
            BoundingBoxXYZ result = null;
            GeometryElement geomElem = famInst
                    .get_Geometry(new Options()
                    {
                        DetailLevel = ViewDetailLevel.Fine,
                    });

            foreach (GeometryInstance inst in geomElem)
            {
                Transform transform = inst.Transform;
                
                GeometryElement instGeomElem = inst.GetInstanceGeometry();
                foreach (GeometryObject obj in instGeomElem)
                {
                    Solid solid = obj as Solid;
                    if (solid != null && solid.Volume != 0)
                    {
                        BoundingBoxXYZ bbox = solid.GetBoundingBox();
                        bbox.Transform = transform;
                        result = new BoundingBoxXYZ()
                        {
                            Max = transform.OfPoint(bbox.Max),
                            Min = transform.OfPoint(bbox.Min),
                        };
                    }
                }
            }

            if (result == null)
                throw new Exception($"Не удалось получить геометрию у элемента с id: {famInst.Id}");

            return result;
        }

        /// <summary>
        /// Подготовка параметров для класса HolesDTO
        /// </summary>
        private static HoleDTO PrepareHoleDTOData(FamilyInstance fi, BoundingBoxXYZ fiBBox, IEnumerable<RevitLinkInstance> linkedModels, double absoluteElevBasePnt)
        {
            string fiName = fi.Symbol.FamilyName;
            
            // Считаю отметки и привязываю HoleDTO
            HoleDTO holesDTO;
            double upMinDist = double.MaxValue;
            double downMinDist = double.MaxValue;
            Floor upFloor = null;
            Floor downFloor = null;
            double downBindElev = 0;
            double rlvDist = 0;

            // Создание расширенного контура, для поиска перекрытий, над/под которыми оно расположено
            BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(new Outline(
                new XYZ(fiBBox.Min.X - 10, fiBBox.Min.Y - 10, fiBBox.Min.Z - 50),
                new XYZ(fiBBox.Max.X + 10, fiBBox.Max.Y + 10, fiBBox.Max.Z + 50)));
            
            foreach (RevitLinkInstance linkedModel in linkedModels)
            {
                // Поиск перекрытий, которые пересекаются с расширенным отверстием
                IEnumerable<Floor> floorColl = new FilteredElementCollector(linkedModel.GetLinkDocument())
                    .OfClass(typeof(Floor))
                    .WherePasses(filter)
                    .Where(fl => fl.Name.StartsWith("00_"))
                    .Cast<Floor>();

                // Обработка пересекающихся перекрытий
                foreach (Floor floor in floorColl)
                {
                    XYZ upPrjPoint = floor.GetVerticalProjectionPoint(fiBBox.Max, FloorFace.Bottom);
                    double upDistance = fiBBox.Max.DistanceTo(upPrjPoint);
                    if (Math.Round(fiBBox.Max.Z, 1) <= Math.Round(upPrjPoint.Z, 1) && upDistance <= upMinDist)
                    {
                        bool checkFloorFaces = CheckFloorFaces(floor, upPrjPoint);
                        if (checkFloorFaces)
                        {
                            upFloor = floor;
                            upMinDist = upDistance;
                        }
                    }

                    XYZ downPrjPoint = floor.GetVerticalProjectionPoint(fiBBox.Min, FloorFace.Top);
                    double downDistance = fiBBox.Min.DistanceTo(downPrjPoint);
                    if (Math.Round(fiBBox.Min.Z, 1) >= Math.Round(downPrjPoint.Z, 1) && downDistance <= downMinDist)
                    {
                        bool checkFloorFaces = CheckFloorFaces(floor, downPrjPoint);
                        if (checkFloorFaces)
                        {
                            downFloor = floor;
                            
                            Level downFloorLevel = (Level)linkedModel.GetLinkDocument().GetElement(downFloor.LevelId);
                            downBindElev = downFloorLevel.Elevation;
                            double floorAboveLevHeight = downFloor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsDouble();

                            if ((downDistance + floorAboveLevHeight) > 0)
                                downMinDist = downDistance + floorAboveLevHeight;
                            else
                            {
                                downMinDist = downDistance;
                                rlvDist = downMinDist;
                            }
                            
                            if (downBindElev < fiBBox.Min.Z)
                                rlvDist = downMinDist;
                            else
                                rlvDist = Math.Round(fiBBox.Min.Z - downBindElev, 2);
                        }
                    }
                }
            }

            if (downFloor is null)
                throw new Exception($"KPLN: Ошибка - экземпляр с id: {fi.Id} невозможно определить основу (уровень, на котором оно расположено)");

            // Корректировка значений под старые семейства (там другая точка вставки)
            if (fiName.ToLower().Contains("tsw") || fiName.ToLower().Contains("trw"))
            {
                double correlDist = Math.Abs((Math.Abs(fiBBox.Max.Z) - Math.Abs(fiBBox.Min.Z)) / 2);
                upMinDist -= correlDist;
                downMinDist += correlDist;
                rlvDist = rlvDist + correlDist - fi.LookupParameter("Расширение границ").AsDouble();
                absoluteElevBasePnt = absoluteElevBasePnt + correlDist - fi.LookupParameter("Расширение границ").AsDouble();
            }

            // Генерация HoleDTO
            if (fiName.ToLower().Contains("прямоуг") || fiName.ToLower().Contains("tsw"))
            {
                holesDTO = new HoleDTO()
                {
                    CurrentHole = fi,
                    UpFloorBinding = upFloor,
                    UpFloorDistance = upMinDist,
                    DownFloorBinding = downFloor,
                    DownFloorDistance = downMinDist,
                    DownBindingElevation = downBindElev,
                    BindingPrefixString = "Низ на отм.",
                    RlvElevation = rlvDist,
                    AbsElevation = absoluteElevBasePnt - Math.Abs(fiBBox.Min.Z),
                };
            }
            else if (fiName.ToLower().Contains("кругл") || fiName.ToLower().Contains("trw"))
            {
                holesDTO = new HoleDTO()
                {
                    CurrentHole = fi,
                    UpFloorBinding = upFloor,
                    UpFloorDistance = upMinDist,
                    DownFloorBinding = downFloor,
                    DownFloorDistance = downMinDist,
                    DownBindingElevation = downBindElev,
                    BindingPrefixString = "Центр на отм.",
                    RlvElevation = rlvDist + Math.Abs((Math.Abs(fiBBox.Max.Z) - Math.Abs(fiBBox.Min.Z)) / 2),
                    AbsElevation = absoluteElevBasePnt - (Math.Abs(fiBBox.Max.Z) + Math.Abs(fiBBox.Min.Z)) / 2,
                };
            }
            else
                throw new Exception($"Не удалось разделить семейство {fiName} на прямоугольное или круглое");

            return holesDTO;
        }

        /// <summary>
        /// Проверка перекрытия на факт пересечения с точкой проецирования
        /// </summary>
        private static bool CheckFloorFaces(Floor floor, XYZ prjPoint)
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
                        //bool intRes = flFace.IsInside(new UV(prjPoint.X, prjPoint.Y));
                        //if (intRes)
                        //{
                        //    return true;
                        //}
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
    }
}
