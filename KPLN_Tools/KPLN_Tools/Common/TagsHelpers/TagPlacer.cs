using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;

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
        private readonly FamilySymbol _tagSymbol;
        private readonly AnnotationCollisionResolver _resolver;

        // Минимальный фрагмент, который маркируем
        private const int MIN_LENGHT_MM = 15;
        
        // Шаг подбора позиции вдоль трубы (в долях от длины сегмента)
        private const int STEPS_ALONG = 11;          // 0.50, 0.45, 0.55, 0.40, 0.60 ...
        private const double PERP_OFFSET_MM = 300.0; // перпендикулярный сдвиг при коллизии
        private const int PERP_TRIES = 4;            // ±300, ±600 ...

        public TagPlacer(Document doc, View view, FamilySymbol tagSymbol, AnnotationCollisionResolver resolver)
        {
            _doc = doc;
            _view = view;
            _tagSymbol = tagSymbol;
            _resolver = resolver;
        }

        public void PlaceForBranches(IEnumerable<Pipe> branchSegments, out int placed, out int skipped)
        {
            placed = 0;
            skipped = 0;

            foreach (Pipe pipe in branchSegments)
            {
                if (TryPlaceOnPipe(pipe))
                    placed++;
                else
                    skipped++;
            }
        }

        private bool TryPlaceOnPipe(Pipe pipe)
        {
            if (!(pipe.Location is LocationCurve lc) || !(lc.Curve is Line line))
                return false; // работаем только с прямыми сегментами

            // Пропускаем вертикальные участки — на плане это точка, марка вдоль трубы не имеет смысла (допуск ~8° от вертикали)
            XYZ dir3d = (line.GetEndPoint(1) - line.GetEndPoint(0)).Normalize();
            if (Math.Abs(dir3d.Z) > 0.99) 
                return false;

            XYZ p0 = ProjectToView(line.GetEndPoint(0));
            XYZ p1 = ProjectToView(line.GetEndPoint(1));
            XYZ dir2d = (p1 - p0);
#if Debug2020 || Revit2020
            double dir2d_lenght = UnitUtils.ConvertFromInternalUnits(dir2d.GetLength(), DisplayUnitType.DUT_MILLIMETERS);
#else
            double dir2d_lenght = UnitUtils.ConvertFromInternalUnits(dir2d.GetLength(), UnitTypeId.Millimeters);
#endif

            if (dir2d_lenght < MIN_LENGHT_MM) 
                return false;
            
            XYZ unit = dir2d.Normalize();

            // Угол вдоль трубы, нормализованный к [-π/2, π/2], чтобы текст не оказался вверх ногами
            double angle = Math.Atan2(unit.Y, unit.X);
            if (angle > Math.PI / 2 + 1e-9) 
                angle -= Math.PI;
            else if (angle < -Math.PI / 2 - 1e-9) 
                angle += Math.PI;

            XYZ midpoint = (p0 + p1) * 0.5;

            IndependentTag tag = IndependentTag.Create(
                _doc, 
                _view.Id, 
                new Reference(pipe), 
                false, 
                TagMode.TM_ADDBY_CATEGORY, 
                TagOrientation.Horizontal, 
                midpoint);

            if (tag == null) 
                return false;


            tag.ChangeTypeId(_tagSymbol.Id);


            RotateTagAroundHead(tag, angle);
            _doc.Regenerate();

            // Перебор позиций вдоль сегмента
            double[] tList = BuildAlongFractions();
            foreach (double t in tList)
            {
                XYZ point = p0 + dir2d * t;
                MoveTagHeadTo(tag, point);
                _doc.Regenerate();

                if (!_resolver.HasCollision(tag))
                    return true;
            }

            // Перпендикулярный сдвиг с выноской
            XYZ perp = new XYZ(-unit.Y, unit.X, 0).Normalize();
#if Debug2020 || Revit2020
            double offsetStep = UnitUtils.ConvertToInternalUnits(PERP_OFFSET_MM, DisplayUnitType.DUT_MILLIMETERS);
#else
            double offsetStep = UnitUtils.ConvertToInternalUnits(PERP_OFFSET_MM, UnitTypeId.Millimeters);
#endif

            try { tag.HasLeader = true; } catch { /* ignore */ }

            for (int i = 1; i <= PERP_TRIES; i++)
            {
                foreach (int sign in new[] { 1, -1 })
                {
                    foreach (double t in tList)
                    {
                        XYZ basePt = p0 + dir2d * t;
                        XYZ shifted = basePt + perp * (offsetStep * i * sign);
                        MoveTagHeadTo(tag, shifted);
                        TrySetLeaderEnd(tag, pipe, basePt);
                        _doc.Regenerate();

                        if (!_resolver.HasCollision(tag))
                            return true;
                    }
                }
            }

            _doc.Delete(tag.Id);
            return false;
        }

        private static double[] BuildAlongFractions()
        {
            // 0.50, 0.45, 0.55, 0.40, 0.60, ..., но не выходим за 0.15..0.85
            var list = new List<double> { 0.50 };
            double step = 0.05;
            for (int i = 1; i <= STEPS_ALONG; i++)
            {
                double down = 0.50 - step * i;
                double up = 0.50 + step * i;
                if (down >= 0.15) list.Add(down);
                if (up <= 0.85) list.Add(up);
            }
            return list.ToArray();
        }

        private XYZ ProjectToView(XYZ p)
        {
            // Для плана — просто обнуляем Z до уровня вида (для коллизий и положения тега
            // важна только XY-плоскость).
            double z = 0;
            try { z = _view.Origin.Z; } catch { z = 0; }
            return new XYZ(p.X, p.Y, z);
        }

        private void MoveTagHeadTo(IndependentTag tag, XYZ target)
        {
            XYZ current = tag.TagHeadPosition;
            XYZ delta = target - current;
            if (delta.GetLength() < 1e-9) 
                return;
            
            ElementTransformUtils.MoveElement(_doc, tag.Id, delta);
        }

        /// <summary>
        /// Поворот вдоль трубы — после Regenerate, чтобы TagHeadPosition был актуальным
        /// </summary>
        private bool RotateTagAroundHead(IndependentTag tag, double angle)
        {
            if (Math.Abs(angle) < 1e-9) 
                return false;

#if !Debug2020 && !Revit2020
            XYZ head = tag.TagHeadPosition;
            Line axis = Line.CreateBound(head, head + XYZ.BasisZ);
            ElementTransformUtils.RotateElement(_doc, tag.Id, axis, angle);
#else
            tag.TagOrientation = TagOrientation.Vertical;
#endif

            return true;
        }

        private void TrySetLeaderEnd(IndependentTag tag, Pipe pipe, XYZ pointOnPipe)
        {
            try
            {
#if !Debug2020 && !Revit2020
                tag.SetLeaderEnd(new Reference(pipe), pointOnPipe);
#else
                tag.LeaderEnd = pointOnPipe; // Revit 2020/2021: одиночный хост
#endif
            }
            catch
            {
                // некоторые конфигурации тегов не поддерживают явное задание конца выноски
            }
        }
    }
}
