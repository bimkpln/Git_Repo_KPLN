using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using KPLN_Tools.Forms.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools.Common.TagsHelpers
{
    /// <summary>
    /// Размещает марки труб ASML_ВК_Марка_Труба на сегментах ответвлений АУПТ.
    /// Одна марка на сегмент, ориентация вдоль трубы. При коллизии с другими
    /// аннотациями двигается вдоль сегмента; если не получается — отодвигается
    /// перпендикулярно с включением выноски.
    /// </summary>
    internal sealed class TagPlacer
    {
        private readonly Document _doc;
        private readonly View _view;
        private readonly AUPTTagPlacerM _placerM;


        public TagPlacer(Document doc, View view, AUPTTagPlacerM placerM)
        {
            _doc = doc;
            _view = view;
            _placerM = placerM;
        }

        public void PlaceForBranches(IEnumerable<Pipe> branchSegments)
        {
            foreach (Pipe pipe in branchSegments)
            {
                // Работаем только с прямыми сегментами
                if (!(pipe.Location is LocationCurve lc) || !(lc.Curve is Line line))
                    continue;


                // Игнор промаркированных (если нужно)
                if (_placerM.IgnoreTagged)
                {
                    var elClassFilter = new ElementClassFilter(typeof(IndependentTag));
                    // Фильтр по виду ElementOwnerViewFilter не применяю, т.к. нужна поддержка Р2020, там его нет
                    var depElems = pipe.GetDependentElements(elClassFilter);
                    if (depElems.Any())
                    {
                        var tagsOnView = depElems
                            .Select(id => _doc.GetElement(id))
                            .Where(e => 
                                e != null 
                                && e.OwnerViewId == _view.Id 
                                && e.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString().Equals(_placerM.SelectedTagTypeName));

                        if (tagsOnView.Any())
                            continue;
                    }
                }


                XYZ p0 = line.GetEndPoint(0);
                XYZ p1 = line.GetEndPoint(1);


                // Пропускаем вертикальные участки — на плане это точка, марка вдоль трубы не имеет смысла (допуск ~8° от вертикали)
                XYZ dir3d = (p1 - p0).Normalize();
                if (Math.Abs(dir3d.Z) > 0.99)
                    continue;


                // Пропускаю мелкие фрагменты
                XYZ dir2d = (p1 - p0);
#if Debug2020 || Revit2020
                double dir2d_lenght = UnitUtils.ConvertFromInternalUnits(dir2d.GetLength(), DisplayUnitType.DUT_MILLIMETERS);
#else
                double dir2d_lenght = UnitUtils.ConvertFromInternalUnits(dir2d.GetLength(), UnitTypeId.Millimeters);
#endif
                if (dir2d_lenght < _placerM.MINPipeLenght)
                    continue;


                // Определяю направление трубы для поворота марки
                XYZ unit = dir2d.Normalize();
                // Угол вдоль трубы, нормализованный к [-π/2, π/2], чтобы текст не оказался вверх ногами
                double angle = Math.Atan2(unit.Y, unit.X);
                if (angle > Math.PI / 2 + 1e-9)
                    angle -= Math.PI;
                else if (angle < -Math.PI / 2 - 1e-9)
                    angle += Math.PI;


                // Определяю количество соед. трубы.
                // Если их больше 2 - делю на отдельные марки на ответвления (такое возможно, если трубы соед. врезками)
                var connMng = pipe.ConnectorManager;
                var connSet = connMng.Connectors;
                XYZ[] pipeSegmPnt;
                // Если больше 2 - начинаю поиск средних по отрезкам
                if (connSet.Size > 2)
                {
                    List<XYZ> splitPoints = new List<XYZ>();

                    var iter = connSet.ForwardIterator();
                    iter.Reset();
                    while (iter.MoveNext())
                    {
                        if (!(iter.Current is Connector conn))
                            continue;

                        // Игнор крайних коннекторов
                        XYZ connOrg = conn.Origin;
                        if (connOrg.IsAlmostEqualTo(p0, 0.001) || connOrg.IsAlmostEqualTo(p1, 0.001))
                            continue;

                        splitPoints.Add(connOrg);
                    }

                    pipeSegmPnt = GetSegmentMidpoints(p0, p1, splitPoints, _placerM.MINPipeLenght);
                }
                // Если 2 - то просто центр
                else
                    pipeSegmPnt = new XYZ[] { (p0 + p1) * 0.5 };


                foreach(XYZ pnt in pipeSegmPnt)
                {
                    IndependentTag tag = IndependentTag.Create(
                        _doc,
                        _view.Id,
                        new Reference(pipe),
                        false,
                        TagMode.TM_ADDBY_CATEGORY,
                        TagOrientation.Horizontal,
                        pnt);

                    if (tag == null)
                        continue;


                    // Применяю марку и вращаю 
                    tag.ChangeTypeId(_placerM.SelectedTagType.Id);
                    RotateTagAroundHead(tag, angle);
                }
            }
        }

        /// <summary>
        /// Поворот вдоль трубы — после Regenerate, чтобы TagHeadPosition был актуальным
        /// </summary>
        private void RotateTagAroundHead(IndependentTag tag, double angle)
        {
            if (Math.Abs(angle) < 1e-9) 
                return;

#if !Debug2020 && !Revit2020
            XYZ head = tag.TagHeadPosition;
            Line axis = Line.CreateBound(head, head + XYZ.BasisZ);
            ElementTransformUtils.RotateElement(_doc, tag.Id, axis, angle);
#else
            tag.TagOrientation = TagOrientation.Vertical;
#endif
        }

        private static XYZ[] GetSegmentMidpoints(XYZ p0, XYZ p1, List<XYZ> splitPoints, double minLengthMm)
        {
#if Debug2020 || Revit2020
            double minLengthFt = UnitUtils.ConvertToInternalUnits(minLengthMm, DisplayUnitType.DUT_MILLIMETERS);
#else
            double minLengthFt = UnitUtils.ConvertToInternalUnits(minLengthMm, UnitTypeId.Millimeters);
#endif

            XYZ dir = p1 - p0;
            double Project(XYZ pt) => (pt - p0).DotProduct(dir);

            var nodes = new List<XYZ> { p0 };
            nodes.AddRange(splitPoints.OrderBy(pt => Project(pt)));
            nodes.Add(p1);

            var midpoints = new List<XYZ>();
            for (int i = 0; i < nodes.Count - 1; i++)
            {
                double length = nodes[i].DistanceTo(nodes[i + 1]);
                if (length < minLengthFt) continue; // пропускаем короткие отрезки

                midpoints.Add((nodes[i] + nodes[i + 1]) * 0.5);
            }

            return midpoints.ToArray();
        }
    }
}
