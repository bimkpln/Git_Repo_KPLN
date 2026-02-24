using Autodesk.Navisworks.Api;
using System;
using System.Globalization;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;


using Application = Autodesk.Navisworks.Api.Application;

namespace KPLN_Quantificator.Common
{
    /// <summary>
    /// Построение точек обзора по комментариям объектов:
    /// - Сканирует дерево модели
    /// - Ищет узлы с DisplayName == "Точка"
    /// - Для каждой такой точки пытается найти "Объект → Комментарии" (у самой точки или у предков)
    /// - Сохраняет координату точки (центр BoundingBox) в словарь (comment -> list of points)
    /// - (ОПЦИОНАЛЬНО) Формирует отладочный отчёт и предлагает сохранить его на рабочем столе
    /// - Создаёт Saved Viewpoints: имя = комментарий, центр = среднее всех точек, ZoomBox по bbox (+1000 мм)
    /// </summary>
    public static class BuildViewpointsByObjectComments
    {
        public static void Build()
        {
            Document doc = Application.MainDocument ?? Application.ActiveDocument;
            if (doc == null || doc.IsClear)
            {
                MessageBox.Show("Документ не открыт.");
                return;
            }

            if (doc.Models == null || doc.Models.Count == 0)
            {
                MessageBox.Show("В документе нет моделей.");
                return;
            }

            var pointsByComment = new Dictionary<string, List<Point3D>>(StringComparer.OrdinalIgnoreCase);

            int models = 0;
            int totalNodes = 0;       // сколько всего ModelItem мы просмотрели
            int pointsFound = 0;      // сколько нашли узлов "Точка"
            int commentsFound = 0;    // сколько раз удалось получить непустой комментарий по найденным точкам
            int pointsCollected = 0;  // сколько точек реально добавили в словарь (успешно получили координату)

            // СБОР ТОЧЕК
            for (int m = 0; m < doc.Models.Count; m++)
            {
                Model model = doc.Models[m];
                if (model == null) continue;

                ModelItem root = model.RootItem;
                if (root == null) continue;

                models++;

                foreach (ModelItem mi in EnumerateSelfAndDescendants(root))
                {
                    totalNodes++;

                    string dn = (mi.DisplayName ?? "").Trim();
                    if (!dn.Equals("Точка", StringComparison.OrdinalIgnoreCase))
                        continue;

                    pointsFound++;

                    string comments = null;

                    ModelItem owner = FindFirstAncestorHavingObjectComments(mi);
                    if (owner != null)
                        comments = ReadObjectComments(owner);

                    if (string.IsNullOrWhiteSpace(comments))
                        continue;

                    commentsFound++;

                    string key = comments.Trim();

                    Point3D center;
                    if (!TryGetModelItemCenter(mi, out center))
                        continue;
                    List<Point3D> list;
                    if (!pointsByComment.TryGetValue(key, out list))
                    {
                        list = new List<Point3D>();
                        pointsByComment[key] = list;
                    }

                    list.Add(center);
                    pointsCollected++;
                }
            }

            // DEBUG
            //ShowDebugResultAndOfferSave(pointsByComment, models, totalNodes, pointsFound, commentsFound, pointsCollected);

            // ПОСТРОЕНИЕ ТОЧЕК ОБЗОРА
            BuildSavedViewpoints(doc, pointsByComment);
        }

        // -------------------------
        // ПОИСК НУЖНЫХ ЭЛЕМЕНТОВ
        // -------------------------

        private static IEnumerable<ModelItem> EnumerateSelfAndDescendants(ModelItem root)
        {
            yield return root;
            foreach (ModelItem d in root.Descendants)
                yield return d;
        }

        private static ModelItem FindFirstAncestorHavingObjectComments(ModelItem start)
        {
            if (start == null) return null;

            if (HasObjectComments(start))
                return start;

            foreach (ModelItem anc in EnumerateAncestorsSafe(start))
                if (HasObjectComments(anc))
                    return anc;

            return null;
        }

        private static IEnumerable<ModelItem> EnumerateAncestorsSafe(ModelItem mi)
        {
            if (mi == null) yield break;

            var ancProp = mi.GetType().GetProperty("Ancestors");
            if (ancProp != null)
            {
                object ancObj = null;
                try { ancObj = ancProp.GetValue(mi, null); } catch { ancObj = null; }

                var en = ancObj as System.Collections.IEnumerable;
                if (en != null)
                {
                    foreach (object a in en)
                    {
                        var ami = a as ModelItem;
                        if (ami != null) yield return ami;
                    }
                    yield break;
                }
            }

            var parentProp = mi.GetType().GetProperty("Parent");
            if (parentProp == null) yield break;

            ModelItem cur = mi;
            while (true)
            {
                ModelItem p = null;
                try { p = parentProp.GetValue(cur, null) as ModelItem; } catch { p = null; }
                if (p == null) yield break;

                yield return p;
                cur = p;
            }
        }

        private static bool HasObjectComments(ModelItem mi)
        {
            if (mi == null) return false;

            foreach (PropertyCategory cat in mi.PropertyCategories)
            {
                string catName = (cat.DisplayName ?? "").Trim();
                if (!catName.Equals("Объект", StringComparison.OrdinalIgnoreCase) &&
                    !catName.Equals("Object", StringComparison.OrdinalIgnoreCase))
                    continue;

                foreach (DataProperty prop in cat.Properties)
                {
                    string propName = (prop.DisplayName ?? "").Trim();
                    if (propName.Equals("Комментарии", StringComparison.OrdinalIgnoreCase) ||
                        propName.Equals("Comments", StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }

            return false;
        }

        private static string ReadObjectComments(ModelItem mi)
        {
            foreach (PropertyCategory cat in mi.PropertyCategories)
            {
                string catName = (cat.DisplayName ?? "").Trim();
                if (!catName.Equals("Объект", StringComparison.OrdinalIgnoreCase) &&
                    !catName.Equals("Object", StringComparison.OrdinalIgnoreCase))
                    continue;

                foreach (DataProperty prop in cat.Properties)
                {
                    string propName = (prop.DisplayName ?? "").Trim();
                    if (propName.Equals("Комментарии", StringComparison.OrdinalIgnoreCase) ||
                        propName.Equals("Comments", StringComparison.OrdinalIgnoreCase))
                    {
                        string s = prop.Value != null ? prop.Value.ToDisplayString() : null;
                        return string.IsNullOrWhiteSpace(s) ? null : s.Trim();
                    }
                }
            }
            return null;
        }

        private static bool TryGetModelItemCenter(ModelItem mi, out Point3D center)
        {
            center = new Point3D(0, 0, 0);

            try
            {
                BoundingBox3D bb = mi.BoundingBox();
                if (bb == null || bb.IsEmpty)
                    return false;

                center = bb.Center;
                return true;
            }
            catch
            {
                return false;
            }
        }

        // -------------------------
        // DEBUG: сохранение текстового отчёта на рабочий стол
        // -------------------------
        private static void ShowDebugResultAndOfferSave(Dictionary<string, List<Point3D>> data,
            int models, int totalNodes, int pointsFound, int commentsFound, int pointsCollected)
        {
            string report = BuildReportText(data, models, totalNodes, pointsFound, commentsFound, pointsCollected);

            var r = MessageBox.Show("Сформирован отладочный отчёт.\n\nСохранить его на рабочем столе?", "KPLN: Viewpoints Debug", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (r != MessageBoxResult.Yes)
                return;

            try
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string fileName = "KPLN_ViewpointsByComments_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".txt";
                string path = Path.Combine(desktop, fileName);

                File.WriteAllText(path, report, Encoding.UTF8);

                MessageBox.Show("Сохранено:\n" + path, "KPLN: OK", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Не удалось сохранить отчёт.\n\n" + ex.GetType().Name + ": " + ex.Message,
                    "KPLN: Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private static string BuildReportText(Dictionary<string, List<Point3D>> data,
            int models, int totalNodes, int pointsFound, int commentsFound, int pointsCollected)
        {
            var sb = new StringBuilder();
            sb.AppendLine("DEBUG создан: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.AppendLine();
            sb.AppendLine(string.Format("Model: {0}", models));
            sb.AppendLine(string.Format("ModelItem: {0}", totalNodes));
            sb.AppendLine(string.Format("Точек с DisplayName=\"Точка\": {0}", pointsFound));
            sb.AppendLine(string.Format("Точек с координатами: {0}", pointsCollected));
            sb.AppendLine(string.Format("Точек с непустыми комментариями: {0}", commentsFound));
            sb.AppendLine(string.Format("Необходимо построить точек обзора (групп): {0}", data.Count));
            sb.AppendLine();
            sb.AppendLine(new string('-', 80));

            foreach (var kvp in data.OrderBy(x => x.Key, StringComparer.CurrentCultureIgnoreCase))
            {
                string comment = kvp.Key;
                List<Point3D> pts = kvp.Value ?? new List<Point3D>();

                sb.AppendLine();
                sb.AppendLine(string.Format("[{0}]  points = {1}", comment, pts.Count));

                for (int i = 0; i < pts.Count; i++)
                {
                    Point3D p = pts[i];
                    sb.AppendLine(string.Format("  #{0}: (X={1}, Y={2}, Z={3})",
                        i + 1, Fmt(p.X), Fmt(p.Y), Fmt(p.Z)));
                }
            }

            return sb.ToString();
        }

        private static string Fmt(double v) { return v.ToString("0.###"); }

        private static string BuildSectionBoxJson(BoundingBox3D bbox)
        {
            // Формат, который понимает View.SetClippingPlanes в режиме Box
            // {"Type":"ClipPlaneSet","Version":1,"OrientedBox":{"Type":"OrientedBox3D","Version":1,"Box":[[min],[max]],"Rotation":[0,0,0]},"Enabled":true}
            // (Rotation = 0,0,0 => осевой (axis-aligned) box)
            var ci = CultureInfo.InvariantCulture;

            string minX = bbox.Min.X.ToString(ci);
            string minY = bbox.Min.Y.ToString(ci);
            string minZ = bbox.Min.Z.ToString(ci);

            string maxX = bbox.Max.X.ToString(ci);
            string maxY = bbox.Max.Y.ToString(ci);
            string maxZ = bbox.Max.Z.ToString(ci);

            return "{"
                + "\"Type\":\"ClipPlaneSet\","
                + "\"Version\":1,"
                + "\"OrientedBox\":{"
                    + "\"Type\":\"OrientedBox3D\","
                    + "\"Version\":1,"
                    + "\"Box\":["
                        + "[" + minX + "," + minY + "," + minZ + "],"
                        + "[" + maxX + "," + maxY + "," + maxZ + "]"
                    + "],"
                    + "\"Rotation\":[0,0,0]"
                + "},"
                + "\"Enabled\":true"
            + "}";
        }

        private static bool ApplySectionBoxToActiveView(Document doc, BoundingBox3D bbox)
        {
            if (doc?.ActiveView == null || bbox == null || bbox.IsEmpty)
                return false;

            string json = BuildSectionBoxJson(bbox);

            // Можно SetClippingPlanes, можно TrySetClippingPlanes - TrySet удобнее, так как вернёт false, если что-то не так
            return doc.ActiveView.TrySetClippingPlanes(json);
        }
















        private static void BuildSavedViewpoints(Document doc, Dictionary<string, List<Point3D>> pointsByComment)
        {
            if (doc == null) return;

            if (pointsByComment == null || pointsByComment.Count == 0)
            {
                MessageBox.Show("Не найдено ни одной группы точек с комментариями.");
                return;
            }

            // +1000 мм (в единицах модели)
            double pad = GetPaddingInModelUnits(doc, 1000.0);

            int created = 0;
            int updated = 0;

            using (Transaction t = doc.BeginTransaction("KPLN: Создание Viewpoints из комментариев"))
            {
                foreach (var kvp in pointsByComment)
                {
                    string rawName = (kvp.Key ?? "").Trim();
                    if (string.IsNullOrWhiteSpace(rawName)) continue;

                    List<Point3D> pts = kvp.Value ?? new List<Point3D>();
                    if (pts.Count == 0) continue;

                    string name = SanitizeViewpointName(rawName);

                    BoundingBox3D bbox = BuildBoundingBoxFromPoints(pts, pad);
                    Viewpoint cam = doc.CurrentViewpoint.CreateCopy();
                    cam.ZoomBox(bbox);
                    cam.PointAt(bbox.Center);
                    try { cam.AlignUp(new UnitVector3D(0, 0, 1)); } catch { }
                    doc.ActiveView.CopyViewpointFrom(cam, ViewChange.JumpCut);
                    ApplySectionBoxToActiveView(doc, bbox);
                    Viewpoint finalVp = doc.ActiveView.CreateViewpointCopy();

                    SavedViewpoint svp = new SavedViewpoint(finalVp)
                    {
                        DisplayName = name
                    };

                    // Если уже существует такой SavedViewpoint — удаляем и создаём заново с тем же именем
                    var existing = FindSavedViewpointByName(doc, name, out GroupItem parentGroup);
                    if (existing != null)
                    {
                        bool removed = doc.SavedViewpoints.Remove(parentGroup, existing);
                        if (removed) updated++;
                    }

                    doc.SavedViewpoints.AddCopy(svp);
                    created++;
                }

                t.Commit();
            }

            MessageBox.Show($"Готово.\n" +
                $"Создано/обновлено точек обзора: {created}\n" +
                $"Заменено старых точек обзора: {updated}");
        }

        // -------------------------
        // BoundingBox по списку точек
        // -------------------------
        private static BoundingBox3D BuildBoundingBoxFromPoints(List<Point3D> pts, double pad)
        {
            double minX = pts[0].X, minY = pts[0].Y, minZ = pts[0].Z;
            double maxX = pts[0].X, maxY = pts[0].Y, maxZ = pts[0].Z;

            for (int i = 1; i < pts.Count; i++)
            {
                var p = pts[i];
                if (p.X < minX) minX = p.X;
                if (p.Y < minY) minY = p.Y;
                if (p.Z < minZ) minZ = p.Z;

                if (p.X > maxX) maxX = p.X;
                if (p.Y > maxY) maxY = p.Y;
                if (p.Z > maxZ) maxZ = p.Z;
            }

            // Если все точки совпали — делаем маленький кубик, иначе bbox может быть "плоский"
            if (Math.Abs(maxX - minX) < 1e-9) { minX -= pad; maxX += pad; }
            if (Math.Abs(maxY - minY) < 1e-9) { minY -= pad; maxY += pad; }
            if (Math.Abs(maxZ - minZ) < 1e-9) { minZ -= pad; maxZ += pad; }

            // Расширяем bbox на pad со всех сторон
            minX -= pad; minY -= pad; minZ -= pad;
            maxX += pad; maxY += pad; maxZ += pad;

            return new BoundingBox3D(
                new Point3D(minX, minY, minZ),
                new Point3D(maxX, maxY, maxZ)
            );
        }

        // -------------------------
        // Конвертация "мм -> units модели"
        // -------------------------
        private static double GetPaddingInModelUnits(Document doc, double padMm)
        {
            Units u = Units.Millimeters;
            try
            {
                var firstModel = doc.Models.FirstOrDefault();
                if (firstModel != null)
                    u = firstModel.Units;
            }
            catch {}

            // Безопасно: свитчим по ToString(), чтобы не зависеть от точных enum-имен в разных версиях
            string us = (u.ToString() ?? "").ToLowerInvariant();

            // Приведение 1000 мм к единицам модели
            if (us.Contains("meter")) return padMm / 1000.0;
            if (us.Contains("centimeter")) return padMm / 10.0;
            if (us.Contains("millimeter")) return padMm;
            if (us.Contains("inch")) return padMm / 25.4;
            if (us.Contains("foot") || us.Contains("feet")) return padMm / 304.8;
            return padMm;
        }

        // -------------------------
        // Поиск SavedViewpoint по имени (в дереве Saved Viewpoints)
        // -------------------------
        private static SavedViewpoint FindSavedViewpointByName(Document doc, string name, out GroupItem parent)
        {
            parent = null;
            if (doc == null || string.IsNullOrWhiteSpace(name)) return null;

            GroupItem root = null;
            try { root = doc.SavedViewpoints.RootItem; } catch { root = null; }

            if (root == null) return null;

            return FindSavedViewpointByNameRecursive(root, null, name.Trim(), out parent);
        }

        private static SavedViewpoint FindSavedViewpointByNameRecursive(GroupItem group, GroupItem parentGroup, string name, out GroupItem foundParent)
        {
            foundParent = null;
            if (group == null) return null;

            foreach (SavedItem child in group.Children)
            {
                if (child == null) continue;

                var svp = child as SavedViewpoint;
                if (svp != null)
                {
                    if (string.Equals((svp.DisplayName ?? "").Trim(), name, StringComparison.OrdinalIgnoreCase))
                    {
                        foundParent = parentGroup; 
                        return svp;
                    }
                }

                var subGroup = child as GroupItem;
                if (subGroup != null)
                {
                    GroupItem fp;
                    var res = FindSavedViewpointByNameRecursive(subGroup, subGroup, name, out fp);
                    if (res != null)
                    {
                        foundParent = fp;
                        return res;
                    }
                }
            }

            return null;
        }

        private static string SanitizeViewpointName(string s)
        {
            s = (s ?? "").Trim();
            if (s.Length == 0) return "Viewpoint";

            s = s.Replace("\r", " ").Replace("\n", " ").Trim();

            if (s.Length > 200) s = s.Substring(0, 200).Trim();

            return s.Length == 0 ? "Viewpoint" : s;
        }
    }
}