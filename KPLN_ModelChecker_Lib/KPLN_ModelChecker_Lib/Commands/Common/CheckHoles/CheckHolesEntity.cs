using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace KPLN_ModelChecker_Lib.Commands.Common.CheckHoles
{
    internal class CheckHolesEntity
    {
        private Transform _currentLinkTransform;
        private BoundingBoxXYZ _candidateBounds;

        public CheckHolesEntity(Element elem)
        {
            CurrentElement = elem;
        }

        public CheckHolesEntity(Element elem, RevitLinkInstance linkInstance) : this(elem)
        {
            CurrentLinkInstance = linkInstance;
        }

        public Element CurrentElement { get; }
        public Solid CurrentSolid { get; set; }
        public List<Solid> CurrentSolids { get; private set; } = new List<Solid>();
        public BoundingBoxXYZ CurrentBBox { get; protected set; }
        public RevitLinkInstance CurrentLinkInstance { get; }

        public Transform CurrentLinkTransform
        {
            get
            {
                if (_currentLinkTransform == null && CurrentLinkInstance != null)
                    _currentLinkTransform = CurrentLinkInstance.GetTotalTransform();
                return _currentLinkTransform;
            }
        }

        /// <summary>Только предварительный отбор; точное пересечение проверяется по Solid.</summary>
        public bool MayIntersect(BoundingBoxXYZ holeBounds)
        {
            if (_candidateBounds == null)
            {
                BoundingBoxXYZ bounds = CurrentElement.get_BoundingBox(null);
                if (bounds == null)
                    return true;
                _candidateBounds = TransformBounds(bounds, CurrentLinkTransform);
            }
            const double tolerance = 1e-6;
            return _candidateBounds.Min.X <= holeBounds.Max.X + tolerance && _candidateBounds.Max.X >= holeBounds.Min.X - tolerance
                && _candidateBounds.Min.Y <= holeBounds.Max.Y + tolerance && _candidateBounds.Max.Y >= holeBounds.Min.Y - tolerance
                && _candidateBounds.Min.Z <= holeBounds.Max.Z + tolerance && _candidateBounds.Max.Z >= holeBounds.Min.Z - tolerance;
        }

        public void SetGeometryData(ViewDetailLevel detailLevel)
        {
            if (CurrentSolid != null && CurrentBBox != null)
                return;

            // Собираем все тела ИОС; для основного тела отверстия сохраняем выбор по максимальному объёму.
            List<Solid> solids = new List<Solid>();
            CollectSolids(CurrentElement.get_Geometry(new Options { DetailLevel = detailLevel }), solids, 0);
            if (solids.Count == 0)
                throw new InvalidOperationException("У элемента нет объёмной геометрии");

            List<Solid> transformed = new List<Solid>();
            Solid largest = null;
            double largestVolume = 0;
            foreach (Solid solid in solids)
            {
                Solid current = CurrentLinkInstance == null ? solid : SolidUtils.CreateTransformed(solid, CurrentLinkTransform);
                double volume = Math.Abs(current.Volume);
                transformed.Add(current);
                if (volume > largestVolume)
                {
                    largestVolume = volume;
                    largest = current;
                }
            }
            if (largest == null)
                throw new InvalidOperationException("Не найдено тело с положительным объёмом");

            BoundingBoxXYZ box = largest.GetBoundingBox() ?? throw new InvalidOperationException("Не удалось получить габарит Solid");
            // Solid уже в координатах основной модели: повторно трансформацию связи не применяем.
            BoundingBoxXYZ worldBounds = TransformBounds(box, null);
            CurrentSolids = transformed;
            CurrentSolid = largest;
            CurrentBBox = worldBounds;
        }

        private static void CollectSolids(GeometryElement geometry, List<Solid> result, int depth)
        {
            if (geometry == null)
                return;
            if (depth > 12)
                throw new InvalidOperationException("Слишком много уровней вложенной геометрии");

            foreach (GeometryObject item in geometry)
            {
                if (item is Solid solid && solid.Faces.Size > 0)
                {
                    // Ошибка чтения объёма передаётся в группированный отчёт, а не замалчивается.
                    if (Math.Abs(solid.Volume) > 1e-12)
                        result.Add(solid);
                }
                else if (item is GeometryInstance instance)
                    CollectSolids(instance.GetInstanceGeometry(), result, depth + 1);
            }
        }

        internal static BoundingBoxXYZ TransformBounds(BoundingBoxXYZ box, Transform link)
        {
            Transform transform = link == null ? box.Transform : link.Multiply(box.Transform);
            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
            for (int i = 0; i < 8; i++)
            {
                XYZ point = transform.OfPoint(new XYZ((i & 1) == 0 ? box.Min.X : box.Max.X,
                    (i & 2) == 0 ? box.Min.Y : box.Max.Y, (i & 4) == 0 ? box.Min.Z : box.Max.Z));
                minX = Math.Min(minX, point.X);
                minY = Math.Min(minY, point.Y);
                minZ = Math.Min(minZ, point.Z);
                maxX = Math.Max(maxX, point.X);
                maxY = Math.Max(maxY, point.Y);
                maxZ = Math.Max(maxZ, point.Z);
            }
            return new BoundingBoxXYZ { Min = new XYZ(minX, minY, minZ), Max = new XYZ(maxX, maxY, maxZ) };
        }
    }
}
