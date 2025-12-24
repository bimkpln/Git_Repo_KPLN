using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows;

using Application = Autodesk.Navisworks.Api.Application;
using GroupItem = Autodesk.Navisworks.Api.GroupItem;

namespace KPLN_Quantificator.Common
{
    public static class GroupClashesByObjectCommentInClashDetective
    {

        public static void RunWithConfirm()
        {
            var r = MessageBox.Show(
                "Сгруппировать ClashResults внутри ClashDetective по \"Объект → Комментарии\"?\n" +
                "Будут созданы группы с именами = значениям `Комментариев` и результаты будут перемещены в них.",
                "KPLN: GroupСlashes по комментариям",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (r != MessageBoxResult.Yes) return;

            var doc = Application.MainDocument ?? Application.ActiveDocument;
            if (doc == null || doc.IsClear)
            {
                MessageBox.Show("Документ не открыт.");
                return;
            }

            var clash = doc.GetClash();
            var dct = clash?.TestsData;
            if (dct == null || dct.Tests == null || dct.Tests.Count == 0)
            {
                MessageBox.Show("Clash Tests не найдены.");
                return;
            }

            int tests = 0, results = 0, groupsCreated = 0, moved = 0, commentFound = 0, pointFound = 0;
            var diag = new StringBuilder();

            var expandedSnapshot = SnapshotExpandedState(dct);

            using (var tr = doc.BeginTransaction("KPLN: GroupClash Results by Object Comments"))
            {
                foreach (var test in dct.Tests.OfType<ClashTest>())
                {
                    tests++;

                    var all = CollectAllLeafResults(test);
                    results += all.Count;

                    var groupByName = new Dictionary<string, ClashResultGroup>(StringComparer.OrdinalIgnoreCase);
                    foreach (var g in test.Children.OfType<ClashResultGroup>())
                    {
                        var name = (g.DisplayName ?? "").Trim();
                        if (!string.IsNullOrWhiteSpace(name) && !groupByName.ContainsKey(name))
                            groupByName[name] = g;
                    }

                    var testGuid = test.Guid;

                    foreach (var cr in all)
                    {
                        var point = FindPointNodeForResult(cr);
                        if (point != null) pointFound++;

                        string comments = null;

                        if (point != null)
                        {
                            var owner = FindFirstAncestorHavingObjectComments(point);
                            if (owner != null) comments = ReadObjectComments(owner);
                        }

                        if (string.IsNullOrWhiteSpace(comments))
                        {
                            var any = FindAnyNodeHavingObjectComments(cr);
                            if (any != null) comments = ReadObjectComments(any);
                        }

                        if (!string.IsNullOrWhiteSpace(comments)) commentFound++;

                        var key = string.IsNullOrWhiteSpace(comments) ? "<ПУСТО>" : comments.Trim();

                        ClashResultGroup targetGroup;
                        if (!groupByName.TryGetValue(key, out targetGroup) || targetGroup == null)
                        {
                            targetGroup = new ClashResultGroup
                            {
                                DisplayName = key,                               
                            };

                            var freshTest = ResolveTestByGuid(dct, testGuid);
                            if (freshTest == null) continue;

                            var okAdd = TestsAddChildSmart(dct, freshTest as GroupItem, targetGroup, diag);
                            if (!okAdd)
                            {
                                tr.Commit();
                                MessageBox.Show(
                                    "Не удалось создать группу в Clash Detective.\n\nDIAG:\n" + diag,
                                    "KPLN: ERROR",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Error);
                                return;
                            }

                            freshTest = ResolveTestByGuid(dct, testGuid);
                            targetGroup = freshTest.Children
                                .OfType<ClashResultGroup>()
                                .FirstOrDefault(g => string.Equals((g.DisplayName ?? "").Trim(), key, StringComparison.OrdinalIgnoreCase));

                            if (targetGroup == null)
                                targetGroup = freshTest.Children.OfType<ClashResultGroup>().LastOrDefault();

                            groupByName[key] = targetGroup;
                            groupsCreated++;
                        }

                        var freshTest2 = ResolveTestByGuid(dct, testGuid);
                        if (freshTest2 == null) continue;

                        GroupItem srcParent;
                        int srcIndex;
                        if (!TryFindParentAndIndex(freshTest2 as GroupItem, cr.Guid, out srcParent, out srcIndex))
                            continue;

                        if (targetGroup != null && srcParent != null && srcParent.Guid == targetGroup.Guid)
                            continue;

                        var dstParent = FindGroupByGuid(freshTest2 as GroupItem, targetGroup.Guid);
                        if (dstParent == null) continue;

                        int dstIndex = dstParent.Children.Count;

                        try
                        {
                            dct.TestsMove(srcParent, srcIndex, dstParent, dstIndex);
                            moved++;
                        }
                        catch (Exception exMove)
                        {
                            diag.AppendLine("[MOVE FAIL] " + exMove.GetType().Name + ": " + exMove.Message);
                        }
                    }
                }

                tr.Commit();
            }

            RestoreExpandedState(dct, expandedSnapshot, defaultExpanded: false);

            MessageBox.Show(
                $"Готово.\n\n" +
                $"Tests: {tests}\n" +
                $"Results: {results}\n" +
                $"Comments found: {commentFound}\n" +
                $"Groups created: {groupsCreated}\n" +
                $"Moved: {moved}\n\n" +
                (diag.Length > 0 ? ("DIAG:\n" + diag) : ""),
                "KPLN: Grouping done");


        }




























        private static Dictionary<Guid, bool> SnapshotExpandedState(DocumentClashTests dct)
        {
            var map = new Dictionary<Guid, bool>();

            foreach (var root in dct.Tests.OfType<GroupItem>())
                SnapshotExpandedStateDeep(root, map);

            return map;
        }

        private static void SnapshotExpandedStateDeep(GroupItem item, Dictionary<Guid, bool> map)
        {
            if (item == null) return;

            map[item.Guid] = GetExpandedState(item);

            foreach (var ch in item.Children.OfType<GroupItem>())
                SnapshotExpandedStateDeep(ch, map);
        }

        private static void RestoreExpandedState(DocumentClashTests dct, Dictionary<Guid, bool> map, bool defaultExpanded = false)
        {
            foreach (var root in dct.Tests.OfType<GroupItem>())
                RestoreExpandedStateDeep(root, map, defaultExpanded);
        }

        private static void RestoreExpandedStateDeep(GroupItem item, Dictionary<Guid, bool> map, bool defaultExpanded)
        {
            if (item == null) return;

            bool expanded;
            if (!map.TryGetValue(item.Guid, out expanded))
                expanded = defaultExpanded;

            SetExpandedState(item, expanded);

            foreach (var ch in item.Children.OfType<GroupItem>())
                RestoreExpandedStateDeep(ch, map, defaultExpanded);
        }

        private static bool GetExpandedState(object o)
        {
            foreach (var name in new[] { "IsExpanded", "Expanded", "IsOpen", "Open" })
            {
                var p = o.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (p != null && p.PropertyType == typeof(bool) && p.CanRead)
                {
                    try { return (bool)p.GetValue(o, null); } catch { }
                }
            }
            return false;
        }

        private static void SetExpandedState(object o, bool value)
        {
            foreach (var name in new[] { "IsExpanded", "Expanded", "IsOpen", "Open" })
            {
                var p = o.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (p != null && p.PropertyType == typeof(bool) && p.CanWrite)
                {
                    try { p.SetValue(o, value, null); return; } catch { }
                }
            }
        }


















        private static bool TestsAddChildSmart(DocumentClashTests dct, GroupItem parent, SavedItem child, StringBuilder diag)
        {
            if (dct == null || parent == null || child == null) return false;

            if (TryInvokeDct(dct, "TestsAddCopy", parent, child, diag)) return true;
            if (TryInvokeDct(dct, "TestsAddCopy", parent, child, parent.Children.Count, diag)) return true;

            if (TryInvokeDct(dct, "TestsAdd", parent, child, diag)) return true;
            if (TryInvokeDct(dct, "TestsAdd", parent, child, parent.Children.Count, diag)) return true;

            if (TryInvokeDct(dct, "TestsInsert", parent, parent.Children.Count, child, diag)) return true;

            try
            {
                parent.Children.Add(child);
                return true;
            }
            catch (Exception ex)
            {
                diag.AppendLine("[ADD FALLBACK FAIL] " + ex.GetType().Name + ": " + ex.Message);
            }

            diag.AppendLine("Нет подходящего метода добавления в DocumentClashTests.");
            diag.AppendLine("Доступные методы Tests*:");
            foreach (var m in dct.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public)
                                 .Where(x => x.Name.StartsWith("Tests", StringComparison.OrdinalIgnoreCase))
                                 .Select(x => x.Name).Distinct().OrderBy(x => x))
            {
                diag.AppendLine(" - " + m);
            }

            return false;
        }

        private static bool TryInvokeDct(DocumentClashTests dct, string methodName, object a1, object a2, StringBuilder diag)
        {
            try
            {
                var m = dct.GetType().GetMethod(methodName,
                    BindingFlags.Instance | BindingFlags.Public,
                    null,
                    new[] { a1.GetType(), a2.GetType() },
                    null);

                if (m == null)
                {
                    m = dct.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public)
                        .FirstOrDefault(x =>
                            x.Name == methodName &&
                            x.GetParameters().Length == 2 &&
                            x.GetParameters()[0].ParameterType.IsAssignableFrom(a1.GetType()) &&
                            x.GetParameters()[1].ParameterType.IsAssignableFrom(a2.GetType()));
                }

                if (m == null) return false;

                m.Invoke(dct, new[] { a1, a2 });
                return true;
            }
            catch (TargetInvocationException tie)
            {
                diag.AppendLine($"[{methodName} FAIL] " + (tie.InnerException?.Message ?? tie.Message));
                return false;
            }
            catch (Exception ex)
            {
                diag.AppendLine($"[{methodName} FAIL] " + ex.GetType().Name + ": " + ex.Message);
                return false;
            }
        }

        private static bool TryInvokeDct(DocumentClashTests dct, string methodName, object a1, object a2, object a3, StringBuilder diag)
        {
            try
            {
                var t1 = a1.GetType();
                var t2 = a2.GetType();
                var t3 = a3.GetType();

                var m = dct.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public)
                    .FirstOrDefault(x =>
                        x.Name == methodName &&
                        x.GetParameters().Length == 3 &&
                        x.GetParameters()[0].ParameterType.IsAssignableFrom(t1) &&
                        x.GetParameters()[1].ParameterType.IsAssignableFrom(t2) &&
                        x.GetParameters()[2].ParameterType.IsAssignableFrom(t3));

                if (m == null) return false;

                m.Invoke(dct, new[] { a1, a2, a3 });
                return true;
            }
            catch (TargetInvocationException tie)
            {
                diag.AppendLine($"[{methodName}(3) FAIL] " + (tie.InnerException?.Message ?? tie.Message));
                return false;
            }
            catch (Exception ex)
            {
                diag.AppendLine($"[{methodName}(3) FAIL] " + ex.GetType().Name + ": " + ex.Message);
                return false;
            }
        }

        private static ClashTest ResolveTestByGuid(DocumentClashTests dct, Guid guid)
        {
            return dct.Tests.OfType<ClashTest>().FirstOrDefault(t => t.Guid == guid);
        }

        private static GroupItem FindGroupByGuid(GroupItem root, Guid guid)
        {
            if (root == null) return null;

            if (root.Guid == guid) return root;

            foreach (var ch in root.Children)
            {
                var g = ch as GroupItem;
                if (g == null) continue;

                var found = FindGroupByGuid(g, guid);
                if (found != null) return found;
            }
            return null;
        }

        private static bool TryFindParentAndIndex(GroupItem root, Guid childGuid, out GroupItem parent, out int index)
        {
            parent = null;
            index = -1;
            if (root == null) return false;

            var children = root.Children.OfType<SavedItem>().ToList();

            for (int i = 0; i < children.Count; i++)
            {
                var si = children[i];

                var cr = si as ClashResult;
                if (cr != null && cr.Guid == childGuid)
                {
                    parent = root;
                    index = i;
                    return true;
                }

                var cg = si as ClashResultGroup;
                if (cg != null)
                {
                    GroupItem subParent;
                    int subIndex;
                    if (TryFindParentAndIndex(cg as GroupItem, childGuid, out subParent, out subIndex))
                    {
                        parent = subParent;
                        index = subIndex;
                        return true;
                    }
                }
            }
            return false;
        }

        private static List<ClashResult> CollectAllLeafResults(ClashTest test)
        {
            var list = new List<ClashResult>();
            if (test?.Children == null) return list;
            CollectDeep(test.Children, list);
            return list;
        }

        private static void CollectDeep(SavedItemCollection parent, List<ClashResult> acc)
        {
            foreach (var it in parent)
            {
                var cr = it as ClashResult;
                if (cr != null)
                {
                    acc.Add(cr);
                    continue;
                }

                var cg = it as ClashResultGroup;
                if (cg != null && cg.Children != null)
                    CollectDeep(cg.Children, acc);
            }
        }

        private static ModelItem FindPointNodeForResult(ClashResult r)
        {
            foreach (var root in EnumerateRootsForResult(r))
            {
                var point = FindFirstNodeByDisplayNameDeep(root, "Точка");
                if (point != null)
                    return point;
            }
            return null;
        }

        private static IEnumerable<ModelItem> EnumerateRootsForResult(ClashResult r)
        {
            foreach (var mi in GetModelItemsFromSelectionSafe(r, 1))
                if (mi != null) yield return mi;

            foreach (var mi in GetModelItemsFromSelectionSafe(r, 2))
                if (mi != null) yield return mi;

            var m1 = r.Item1 as ModelItem;
            if (m1 != null) yield return m1;

            var m2 = r.Item2 as ModelItem;
            if (m2 != null) yield return m2;
        }

        private static ModelItem FindFirstNodeByDisplayNameDeep(ModelItem root, string displayName)
        {
            if (root == null) return null;

            foreach (var mi in EnumerateSelfAndDescendants(root))
            {
                var dn = (mi.DisplayName ?? "").Trim();
                if (dn.Equals(displayName, StringComparison.OrdinalIgnoreCase))
                    return mi;
            }
            return null;
        }

        private static IEnumerable<ModelItem> EnumerateSelfAndDescendants(ModelItem root)
        {
            yield return root;
            foreach (var d in root.Descendants)
                yield return d;
        }

        private static ModelItem FindFirstAncestorHavingObjectComments(ModelItem start)
        {
            if (start == null) return null;

            if (HasObjectComments(start)) return start;

            foreach (var anc in EnumerateAncestorsSafe(start))
                if (HasObjectComments(anc)) return anc;

            return null;
        }

        private static ModelItem FindAnyNodeHavingObjectComments(ClashResult r)
        {
            foreach (var root in EnumerateRootsForResult(r))
            {
                foreach (var mi in EnumerateSelfAndDescendants(root))
                {
                    if (HasObjectComments(mi))
                        return mi;

                    foreach (var anc in EnumerateAncestorsSafe(mi))
                        if (HasObjectComments(anc)) return anc;
                }
            }
            return null;
        }

        private static bool HasObjectComments(ModelItem mi)
        {
            if (mi == null) return false;

            foreach (var cat in mi.PropertyCategories)
            {
                var catName = (cat.DisplayName ?? "").Trim();
                if (!catName.Equals("Объект", StringComparison.OrdinalIgnoreCase) &&
                    !catName.Equals("Object", StringComparison.OrdinalIgnoreCase))
                    continue;

                foreach (var prop in cat.Properties)
                {
                    var propName = (prop.DisplayName ?? "").Trim();
                    if (propName.Equals("Комментарии", StringComparison.OrdinalIgnoreCase) ||
                        propName.Equals("Comments", StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            return false;
        }

        private static string ReadObjectComments(ModelItem mi)
        {
            foreach (var cat in mi.PropertyCategories)
            {
                var catName = (cat.DisplayName ?? "").Trim();
                if (!catName.Equals("Объект", StringComparison.OrdinalIgnoreCase) &&
                    !catName.Equals("Object", StringComparison.OrdinalIgnoreCase))
                    continue;

                foreach (var prop in cat.Properties)
                {
                    var propName = (prop.DisplayName ?? "").Trim();
                    if (propName.Equals("Комментарии", StringComparison.OrdinalIgnoreCase) ||
                        propName.Equals("Comments", StringComparison.OrdinalIgnoreCase))
                    {
                        var s = prop.Value != null ? prop.Value.ToDisplayString() : null;
                        return string.IsNullOrWhiteSpace(s) ? null : s.Trim();
                    }
                }
            }
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
                    foreach (var a in en)
                        if (a is ModelItem ami) yield return ami;

                    yield break;
                }
            }

            var parentProp = mi.GetType().GetProperty("Parent");
            if (parentProp == null) yield break;

            var cur = mi;
            while (true)
            {
                ModelItem p = null;
                try { p = parentProp.GetValue(cur, null) as ModelItem; } catch { p = null; }
                if (p == null) yield break;

                yield return p;
                cur = p;
            }
        }

        private static IEnumerable<ModelItem> GetModelItemsFromSelectionSafe(ClashResult r, int side)
        {
            object selObj = null;
            try
            {
                selObj = (side == 1)
                    ? r.GetType().GetProperty("Selection1")?.GetValue(r, null)
                    : r.GetType().GetProperty("Selection2")?.GetValue(r, null);
            }
            catch { yield break; }

            if (selObj == null) yield break;

            object selectedItemsObj = null;
            try
            {
                selectedItemsObj = selObj.GetType().GetProperty("SelectedItems")?.GetValue(selObj, null);
            }
            catch { yield break; }

            var mic = selectedItemsObj as ModelItemCollection;
            if (mic != null)
            {
                foreach (ModelItem mi in mic)
                    if (mi != null) yield return mi;
            }
        }
    }
}