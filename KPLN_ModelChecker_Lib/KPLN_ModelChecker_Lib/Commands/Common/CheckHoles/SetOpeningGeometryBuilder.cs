using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace KPLN_ModelChecker_Lib.Commands.Common.CheckHoles
{
    /// <summary>Расчётное тело отверстия СЕТ без Solid — по двум модельным контурам семейства.</summary>
    internal static class SetOpeningGeometryBuilder
    {
        private const double _tolerance = 0.001 / 304.8;

        internal static Solid Create(FamilyInstance hole, GeometryElement geometry)
        {
            if (!CheckHolesProjectRules.IsSetWallOpening(hole.Document.Title, hole.Symbol.FamilyName,
                hole.Category != null && hole.Category.Id.Equals(new ElementId(BuiltInCategory.OST_Windows))))
                throw new InvalidOperationException("Построение по модельным линиям разрешено только для отверстий СЕТ");

            // Читаем линии экземпляра, а не общий габарит с условным оформлением.
            Transform transform = hole.GetTransform();
            List<Line> lines = new List<Line>();
            CollectLines(hole.Document, geometry, transform.Inverse, lines, 0);

            return BuildSolid(lines, transform);
        }

        private static void CollectLines(Document document, GeometryElement geometry, Transform inverse,
            List<Line> lines, int depth)
        {
            if (geometry == null)
                return;
            if (depth > 12)
                throw new InvalidOperationException("Слишком много уровней вложенных модельных линий");

            foreach (GeometryObject item in geometry)
            {
                if (item is GeometryInstance instance)
                {
                    CollectLines(document, instance.GetInstanceGeometry(), inverse, lines, depth + 1);
                    continue;
                }

                if (!(item is Curve curve))
                    continue;

                GraphicsStyle style = document.GetElement(curve.GraphicsStyleId) as GraphicsStyle;
                Category category = style?.GraphicsStyleCategory;
                // Выбранные в семействе контуры относятся к «Проем» и основной категории «Окна».
                if (category == null || (!category.Id.Equals(new ElementId(BuiltInCategory.OST_Windows))
                    && !string.Equals(category.Name, "Проем", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(category.Name, "Проём", StringComparison.OrdinalIgnoreCase)))
                    continue;

                if (!(curve is Line) || !curve.IsBound)
                    throw new InvalidOperationException("Контуры отверстия содержат кривые вместо прямых отрезков");

                lines.Add(Line.CreateBound(inverse.OfPoint(curve.GetEndPoint(0)), inverse.OfPoint(curve.GetEndPoint(1))));
            }
        }

        private static Solid BuildSolid(List<Line> lines, Transform transform)
        {
            if (lines.Count != 8)
                throw new InvalidOperationException("Ожидались восемь модельных линий двух прямоугольных контуров; найдено: " + lines.Count);

            // В локальных координатах семейства проход направлен по Y, сечение лежит в XZ.
            List<Line> first = new List<Line>();
            List<Line> second = new List<Line>();
            double firstY = lines[0].GetEndPoint(0).Y;
            double? secondY = null;
            foreach (Line line in lines)
            {
                XYZ start = line.GetEndPoint(0);
                XYZ end = line.GetEndPoint(1);
                if (!Near(start.Y, end.Y))
                    throw new InvalidOperationException("Линии контура не лежат в плоскости поперёк отверстия");

                if (Near(start.Y, firstY))
                    first.Add(line);
                else
                {
                    if (!secondY.HasValue)
                        secondY = start.Y;
                    if (!Near(start.Y, secondY.Value))
                        throw new InvalidOperationException("Найдено более двух плоскостей контуров отверстия");

                    second.Add(line);
                }
            }

            if (!secondY.HasValue || Math.Abs(firstY - secondY.Value) <= _tolerance)
                throw new InvalidOperationException("Не удалось определить глубину отверстия между контурами");

            Rectangle front = ReadRectangle(first);
            Rectangle back = ReadRectangle(second);
            if (!Near(front.MinX, back.MinX) || !Near(front.MaxX, back.MaxX)
                || !Near(front.MinZ, back.MinZ) || !Near(front.MaxZ, back.MaxZ))
                throw new InvalidOperationException("Контуры отверстия не совпадают по размерам или положению");

            // Глубина берётся именно между выбранными модельными контурами, включая их отступы.
            double nearY = Math.Min(firstY, secondY.Value);
            double depth = Math.Abs(firstY - secondY.Value);
            XYZ[] points =
            {
                new XYZ(front.MinX, nearY, front.MinZ),
                new XYZ(front.MaxX, nearY, front.MinZ),
                new XYZ(front.MaxX, nearY, front.MaxZ),
                new XYZ(front.MinX, nearY, front.MaxZ)
            };
            CurveLoop loop = new CurveLoop();
            for (int i = 0; i < points.Length; i++)
                loop.Append(Line.CreateBound(points[i], points[(i + 1) % points.Length]));

            using (Solid local = GeometryCreationUtilities.CreateExtrusionGeometry(new[] { loop }, XYZ.BasisY, depth))
            {
                // Возвращаем тело в координатах экземпляра; в документ ничего не добавляем.
                return SolidUtils.CreateTransformed(local, transform);
            }
        }

        private static Rectangle ReadRectangle(List<Line> lines)
        {
            if (lines.Count != 4)
                throw new InvalidOperationException("Каждый контур должен состоять из четырёх модельных линий");

            Rectangle result = new Rectangle();
            foreach (Line line in lines)
            {
                for (int end = 0; end < 2; end++)
                {
                    XYZ point = line.GetEndPoint(end);
                    result.MinX = Math.Min(result.MinX, point.X);
                    result.MaxX = Math.Max(result.MaxX, point.X);
                    result.MinZ = Math.Min(result.MinZ, point.Z);
                    result.MaxZ = Math.Max(result.MaxZ, point.Z);
                }
            }

            if (result.MaxX - result.MinX <= _tolerance || result.MaxZ - result.MinZ <= _tolerance)
                throw new InvalidOperationException("Контур отверстия имеет нулевую ширину или высоту");

            // Одних габаритов недостаточно: подтверждаем все четыре стороны без разрывов и дублей.
            int edges = 0;
            foreach (Line line in lines)
            {
                int start = result.Corner(line.GetEndPoint(0));
                int end = result.Corner(line.GetEndPoint(1));
                int edge = -1;
                for (int i = 0; i < 4; i++)
                {
                    int next = (i + 1) % 4;
                    if ((start == i && end == next) || (end == i && start == next))
                        edge = i;
                }

                if (edge < 0 || (edges & (1 << edge)) != 0)
                    throw new InvalidOperationException("Модельные линии не образуют замкнутый прямоугольник без повторяющихся сторон");

                edges |= 1 << edge;
            }

            return result;
        }

        private static bool Near(double first, double second) => Math.Abs(first - second) <= _tolerance;

        private sealed class Rectangle
        {
            internal double MinX = double.MaxValue;
            internal double MaxX = double.MinValue;
            internal double MinZ = double.MaxValue;
            internal double MaxZ = double.MinValue;

            internal int Corner(XYZ point)
            {
                if (Near(point.X, MinX) && Near(point.Z, MinZ))
                    return 0;
                if (Near(point.X, MaxX) && Near(point.Z, MinZ))
                    return 1;
                if (Near(point.X, MaxX) && Near(point.Z, MaxZ))
                    return 2;
                if (Near(point.X, MinX) && Near(point.Z, MaxZ))
                    return 3;

                return -1;
            }
        }
    }
}
