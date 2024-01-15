using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Tools.Common.HolesManager
{
    /// <summary>
    /// Класс для обработки спец. класса IOSShaftDTO
    /// </summary>
    internal class IOSShaftPrepareManager
    {
        private readonly IEnumerable<FamilyInstance> _shaftElems;
        private readonly IEnumerable<RevitLinkInstance> _linkedModels;

        public IOSShaftPrepareManager(IEnumerable<FamilyInstance> shaftElems, IEnumerable<RevitLinkInstance> linkedModels)
        {
            _shaftElems = shaftElems;
            _linkedModels = linkedModels;
        }

        /// <summary>
        /// Коллекция ошибок, при генерации IOSShaftDTO
        /// </summary>
        public List<FamilyInstance> ErrorFamInstColl { get; private set; } = new List<FamilyInstance>();

        /// <summary>
        /// Подготовка спец. семейст для анализа
        /// </summary>
        public List<IOSShaftDTO> PrepareShaftDTO()
        {
            List<IOSShaftDTO> result = new List<IOSShaftDTO>();
            foreach (FamilyInstance fi in _shaftElems)
            {
                BoundingBoxXYZ fiBBox = PrepareHoleBBox(fi);
                IOSShaftDTO shaftDTO = PrepareHoleDTOData(fi, fiBBox);
                if (shaftDTO != null)
                    result.Add(shaftDTO);
            }

            return result;
        }

        /// <summary>
        /// Подготовка геометрии элемента отверстия
        /// </summary>
        private BoundingBoxXYZ PrepareHoleBBox(FamilyInstance famInst)
        {
            BoundingBoxXYZ fiBBox = famInst
                        .get_Geometry(new Options()
                        {
                            DetailLevel = ViewDetailLevel.Fine,
                        })
                        .GetBoundingBox();

            if (fiBBox == null)
                throw new Exception($"Не удалось получить геометрию у элемента с id: {famInst.Id}");
            
            return fiBBox;
        }

        /// <summary>
        /// Подготовка параметров для класса HolesDTO
        /// </summary>
        private IOSShaftDTO PrepareHoleDTOData(FamilyInstance fi, BoundingBoxXYZ fiBBox)
        {
            string fiName = fi.Symbol.FamilyName;

            // Считаю отметки и привязываю ShaftDTO
            IOSShaftDTO shaftDTO = null;
            Element downHost = null;
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
                List<Element> arElemColl = new List<Element>();

                arElemColl.AddRange(new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(Floor))
                    .WherePasses(filter)
                    .Where(fl => fl.Name.StartsWith("00_") || fl.Name.StartsWith("KTS_00_"))
                    .Cast<Floor>());
                arElemColl.AddRange(new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(RoofBase))
                    .WherePasses(filter)
                    .Where(fl => fl.Name.StartsWith("01_"))
                    .Cast<RoofBase>());

                // Обработка пересекающихся перекрытий
                if (arElemColl.Any())
                {
                    downBindElev = arElemColl.Min(e => ((Level)linkDoc.GetElement(e.LevelId)).Elevation);
                    downHost = arElemColl.Where(e => ((Level)linkDoc.GetElement(e.LevelId)).Elevation == downBindElev).FirstOrDefault();
                    rlvDist = fiBBox.Min.Z - downBindElev;
                }
            }

            if (downHost is null)
            {
                ErrorFamInstColl.Add(fi);
                return null;
            }

            shaftDTO = new IOSShaftDTO()
            {
                CurrentHole = fi,
                DownFloorBinding = downHost,
                DownBindingElevation = downBindElev,
                BindingPrefixString = "Низ на отм.",
                RlvElevation = rlvDist,
                AbsElevation = fiBBox.Min.Z,
            };

            return shaftDTO;
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
                transform.Inverse.OfPoint(new XYZ(sminX + 0.5, sminY + 0.5, bbox.Min.Z - 1)),
                transform.Inverse.OfPoint(new XYZ(smaxX + 0.5, smaxY + 0.5, bbox.Max.Z + 1)));

            return new BoundingBoxIntersectsFilter(outline);
        }
    }
}
