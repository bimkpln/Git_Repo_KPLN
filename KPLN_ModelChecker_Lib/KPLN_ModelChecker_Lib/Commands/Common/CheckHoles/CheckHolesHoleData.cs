using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace KPLN_ModelChecker_Lib.Commands.Common.CheckHoles
{
    internal sealed class CheckHolesHoleData : CheckHolesEntity
    {
        public double? MainHoleArea { get; private set; }
        public Dictionary<string, HashSet<ElementId>> FailedElements { get; } = new Dictionary<string, HashSet<ElementId>>();
        public List<string> GeometryReadErrors { get; } = new List<string>();
        public List<CheckHolesMEPData> IntesectElementsColl { get; } = new List<CheckHolesMEPData>();

        /// <summary>Площадь объединённой проекции, кв. футы; без повторного учёта наложенных областей.</summary>
        public double SumIntersectArea { get; private set; }

        public CheckHolesHoleData(Element elem) : base(elem) { }

        private void AddGeometryError(CheckHolesMEPData data, string error, Action<CheckHolesEntity, string> reportError)
        {
            reportError(data, error);
            string source = $"{data.CurrentElement.Document.Title} (ID связи: {data.CurrentLinkInstance?.Id})";
            if (!FailedElements.TryGetValue(source, out HashSet<ElementId> ids))
            {
                ids = new HashSet<ElementId>();
                FailedElements.Add(source, ids);
            }
            ids.Add(data.CurrentElement.Id);
            GeometryReadErrors.Add($"{source}; элемент ИОС: {data.CurrentElement.Id}. {error}");
        }

        public void SetIntersectsData(List<CheckHolesMEPData> mepElements, Action<CheckHolesEntity, string> reportError)
        {
            HoleProjection projection = new HoleProjection((FamilyInstance)CurrentElement, CurrentSolid);
            MainHoleArea = projection.OpeningArea;
            foreach (CheckHolesMEPData data in mepElements)
            {
                try
                {
                    // Проверяем габарит именно этого отверстия; близость точки вставки не доказывает пересечение.
                    if (!data.MayIntersect(CurrentBBox))
                        continue;
                }
                catch (Exception ex)
                {
                    AddGeometryError(data, $"Предварительный отбор: {ex.GetType().FullName}: {ex.Message}", reportError);
                    continue;
                }

                if (data.GeometryReadError != null)
                {
                    AddGeometryError(data, data.GeometryReadError, reportError);
                    continue;
                }
                try
                {
                    data.SetGeometryData(ViewDetailLevel.Fine);
                }
                catch (Exception ex)
                {
                    data.CurrentSolid = null;
                    data.GeometryReadError = $"Чтение геометрии ИОС: {ex.GetType().FullName}: {ex.Message}";
                    AddGeometryError(data, data.GeometryReadError, reportError);
                    continue;
                }

                List<Solid> intersections = new List<Solid>();
                try
                {
                    // Обрезаем по глубине; внешний контур затем ограничивается плоскостью проёма.
                    foreach (Solid solid in data.CurrentSolids)
                    {
                        Solid intersection = projection.ClipToDepth(solid);
                        if (intersection == null)
                            continue;
                        intersections.Add(intersection);
                        if (intersection.Faces.Size == 0 || Math.Abs(intersection.Volume) <= 1e-12)
                        {
                            intersections.RemoveAt(intersections.Count - 1);
                            intersection.Dispose();
                        }
                    }

                    if (intersections.Count > 0 && projection.Add(intersections))
                    {
                        IntesectElementsColl.Add(data);
                        SumIntersectArea = projection.OccupiedArea;
                    }
                }
                catch (Exception ex)
                {
                    AddGeometryError(data, $"Расчёт проекции пересечения: {ex.GetType().FullName}: {ex.Message}", reportError);
                }
                finally
                {
                    foreach (Solid intersection in intersections)
                        intersection.Dispose();
                }
            }
        }
    }
}
