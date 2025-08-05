﻿using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static KPLN_Quantificator.Common.Collections;
using Application = Autodesk.Navisworks.Api.Application;

namespace KPLN_Quantificator
{
    public class GroupingFunctions
    {
        public static void GroupClashes(ClashTest selectedClashTest, GroupingMode groupingMode, GroupingMode subgroupingMode, bool keepExistingGroups, string clashStatus)
        {
            try
            {
                // Шаг 1: Получаем все конфликты (включая те, что в группах)
                List<ClashResult> allClashResults = GetIndividualClashResults(selectedClashTest, true).ToList();

                // Шаг 2: Фильтрация по статусу
                List<ClashResult> filteredClashResults = allClashResults.Where(cr =>
                {
                    string status = cr.Status.ToString();

                    switch (clashStatus)
                    {
                        case "NewClash":
                            return status == "New";
                        case "ActiveClash":
                            return status == "Active";
                        case "NewActiveClash":
                            return status == "New" || status == "Active";
                        case "AllClash":
                            return true;
                        default:
                            return true;
                    }
                }).ToList();

                var filteredGuids = filteredClashResults.Select(cr => cr.Guid).ToHashSet();

                // Шаг 3: Группировка только отфильтрованных
                List<ClashResultGroup> clashResultGroups = new List<ClashResultGroup>();
                CreateGroup(ref clashResultGroups, groupingMode, filteredClashResults, "");

                if (subgroupingMode != GroupingMode.None)
                    CreateSubGroups(ref clashResultGroups, subgroupingMode);

                // Шаг 4: Удаление групп с 1 конфликтом
                List<ClashResult> ungroupedClashResults = RemoveOneClashGroup(ref clashResultGroups);

                // Шаг 5: Конфликты вне фильтра
                List<ClashResult> untouchedClashResults = allClashResults
                    .Where(cr => !filteredGuids.Contains(cr.Guid))
                    .ToList();

                // Шаг 6: Группы вне фильтра и внутри фильтра (если нужно)
                List<ClashResultGroup> allExistingGroups = BackupExistingClashGroups(selectedClashTest).ToList();

                List<ClashResultGroup> untouchedGroups = allExistingGroups
                    .Where(gr => gr.Children.OfType<ClashResult>().All(cr => !filteredGuids.Contains(cr.Guid)))
                    .ToList();

                List<ClashResultGroup> oldGroupsToPreserve = allExistingGroups
                    .Where(gr => gr.Children.OfType<ClashResult>().Any(cr => filteredGuids.Contains(cr.Guid)))
                    .ToList();

                if (keepExistingGroups)
                    clashResultGroups.AddRange(oldGroupsToPreserve);

                // ❗ Шаг 7: Создаём копии всех элементов ДО замены теста
                var clashGroupsCopy = clashResultGroups
                    .Select(gr => (ClashResultGroup)gr.CreateCopy())
                    .ToList();

                var ungroupedResultsCopy = ungroupedClashResults
                    .Select(cr => (ClashResult)cr.CreateCopy())
                    .ToList();

                var untouchedGroupsCopy = untouchedGroups
                    .Select(gr => (ClashResultGroup)gr.CreateCopy())
                    .ToList();

                var untouchedResultsCopy = untouchedClashResults
                    .Select(cr => (ClashResult)cr.CreateCopy())
                    .ToList();

                // Шаг 8: Записываем в новый тест
                ProcessClashGroupPreservingOthers(
                    clashGroupsCopy,
                    ungroupedResultsCopy,
                    untouchedGroupsCopy,
                    untouchedResultsCopy,
                    selectedClashTest
                );
            }
            catch (Exception ex)
            {
                // Показываем ошибку в окне
                System.Windows.Forms.MessageBox.Show(ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }




        private static void ProcessClashGroupPreservingOthers(List<ClashResultGroup> clashGroups,List<ClashResult> ungroupedClashResults,List<ClashResultGroup> untouchedGroups,
            List<ClashResult> untouchedClashResults, ClashTest selectedClashTest)
        {
            try
            {
                using (Transaction tx = Application.MainDocument.BeginTransaction("Group clashes"))
                {
                    ClashTest copiedClashTest = (ClashTest)selectedClashTest.CreateCopyWithoutChildren();
                    ClashTest backupTest = (ClashTest)selectedClashTest.CreateCopy();

                    DocumentClash documentClash = Application.MainDocument.GetClash();
                    int indexOfClashTest = documentClash.TestsData.Tests.IndexOf(selectedClashTest);

                    // Заменяем тест на копию без детей
                    documentClash.TestsData.TestsReplaceWithCopy(indexOfClashTest, copiedClashTest);

                    int totalItems = clashGroups.Count + ungroupedClashResults.Count + untouchedGroups.Count + untouchedClashResults.Count;
                    int currentProgress = 0;

                    Progress progressBar = Application.BeginProgress(
                        "Copying Results",
                        $"Copying results from {selectedClashTest.DisplayName} to the Group Clashes pane...");

                    GroupItem groupItem = (GroupItem)documentClash.TestsData.Tests[indexOfClashTest];

                    // Новые группы
                    foreach (var clashGroup in clashGroups)
                    {
                        if (progressBar.IsCanceled) break;

                        var copy = (SavedItem)clashGroup.CreateCopy();
                        documentClash.TestsData.TestsAddCopy(groupItem, copy);

                        currentProgress++;
                        progressBar.Update((double)currentProgress / totalItems);
                    }

                    // Отдельные конфликты из фильтра
                    foreach (var clash in ungroupedClashResults)
                    {
                        if (progressBar.IsCanceled) break;

                        var copy = (SavedItem)clash.CreateCopy();
                        documentClash.TestsData.TestsAddCopy(groupItem, copy);

                        currentProgress++;
                        progressBar.Update((double)currentProgress / totalItems);
                    }

                    // СТАРЫЕ ГРУППЫ вне фильтра
                    foreach (var group in untouchedGroups)
                    {
                        if (progressBar.IsCanceled) break;

                        var copy = (SavedItem)group.CreateCopy();
                        documentClash.TestsData.TestsAddCopy(groupItem, copy);

                        currentProgress++;
                        progressBar.Update((double)currentProgress / totalItems);
                    }

                    // СТАРЫЕ отдельные конфликты вне фильтра
                    foreach (var clash in untouchedClashResults)
                    {
                        if (progressBar.IsCanceled) break;

                        var copy = (SavedItem)clash.CreateCopy();
                        documentClash.TestsData.TestsAddCopy(groupItem, copy);

                        currentProgress++;
                        progressBar.Update((double)currentProgress / totalItems);
                    }

                    // Если отменили — откатываем
                    if (progressBar.IsCanceled)
                    {
                        documentClash.TestsData.TestsReplaceWithCopy(indexOfClashTest, backupTest);
                    }

                    tx.Commit();
                    Application.EndProgress();
                }
            }
            catch (Exception ex) 
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }















        private static void CreateGroup(ref List<ClashResultGroup> clashResultGroups, GroupingMode groupingMode, List<ClashResult> clashResults, string initialName)
        {
            switch (groupingMode)
            {
                case GroupingMode.None:
                    return;
                case GroupingMode.Level:
                    clashResultGroups = GroupByLevel(clashResults, initialName);
                    break;
                case GroupingMode.GridIntersection:
                    clashResultGroups = GroupByGridIntersection(clashResults, initialName);
                    break;
                case GroupingMode.SelectionA:
                case GroupingMode.SelectionB:
                    clashResultGroups = GroupByElementOfAGivenSelection(clashResults, groupingMode, initialName);
                    break;
                case GroupingMode.ApprovedBy:
                case GroupingMode.AssignedTo:
                case GroupingMode.Status:
                    clashResultGroups = GroupByProperties(clashResults, groupingMode, initialName);
                    break;
                case GroupingMode.Host:
                    clashResultGroups = GroupByHost(clashResults);
                    break;
            }
        }

        private static void CreateSubGroups(ref List<ClashResultGroup> clashResultGroups, GroupingMode mode)
        {
            List<ClashResultGroup> clashResultSubGroups = new List<ClashResultGroup>();

            foreach (ClashResultGroup group in clashResultGroups)
            {
                List<ClashResult> clashResults = new List<ClashResult>();

                foreach (SavedItem item in group.Children)
                {
                    ClashResult clashResult = item as ClashResult;
                    if (clashResult != null)
                    {
                        clashResults.Add(clashResult);
                    }
                }

                List<ClashResultGroup> clashResultTempSubGroups = new List<ClashResultGroup>();
                CreateGroup(ref clashResultTempSubGroups, mode, clashResults, group.DisplayName + "_");
                clashResultSubGroups.AddRange(clashResultTempSubGroups);
            }
            clashResultGroups = clashResultSubGroups;
        }

        public static void UnGroupClashes(ClashTest selectedClashTest)
        {
            List<ClashResultGroup> groups = new List<ClashResultGroup>();
            List<ClashResult> results = GetIndividualClashResults(selectedClashTest, false).ToList();
            List<ClashResult> copiedResult = new List<ClashResult>();

            foreach (ClashResult result in results)
            {
                copiedResult.Add((ClashResult)result.CreateCopy());
            }

            ProcessClashGroup(groups, copiedResult, selectedClashTest);

        }

        #region grouping functions
        private static int GetGroupSize(List<ClashResult> results, ModelItem item)
        {
            if (item == null) { return 0; }
            int ammount = 0;
            foreach (ClashResult result in results)
            {
                ClashResult copiedResult = (ClashResult)result.CreateCopy();
                if (copiedResult.CompositeItem1 != null)
                {
                    ModelItem modelItem = GetSignificantAncestorOrSelf(copiedResult.CompositeItem1);
                    if (modelItem != null)
                    {
                        if (modelItem == item)
                        {
                            ammount++;
                        }
                    }
                }
                if (copiedResult.CompositeItem2 != null)
                {
                    ModelItem modelItem = GetSignificantAncestorOrSelf(copiedResult.CompositeItem2);
                    if (modelItem != null)
                    {
                        if (modelItem == item)
                        {
                            ammount++;
                        }
                    }
                }
            }
            return ammount;
        }
        private static List<ClashResultGroup> GroupByHost(List<ClashResult> results)
        {
            Dictionary<ModelItem, ClashResultGroup> groups = new Dictionary<ModelItem, ClashResultGroup>();
            ClashResultGroup currentGroup;
            List<ClashResultGroup> emptyClashResultGroups = new List<ClashResultGroup>();
            int num = 0;
            foreach (ClashResult result in results)
            {
                //Cannot add original result to new clash test, so I create a copy
                ClashResult copiedResult = (ClashResult)result.CreateCopy();
                ModelItem modelItem = null;
                int s = 0;
                int s1 = GetGroupSize(results, GetSignificantAncestorOrSelf(copiedResult.CompositeItem1));
                int s2 = GetGroupSize(results, GetSignificantAncestorOrSelf(copiedResult.CompositeItem2));
                if (s1 >= s2)
                {
                    s = s1;
                    modelItem = GetSignificantAncestorOrSelf(copiedResult.CompositeItem1);
                }
                else
                {
                    s = s2;
                    modelItem = GetSignificantAncestorOrSelf(copiedResult.CompositeItem2);
                }
                string displayName = "Пусто";
                if (modelItem != null)
                {
                    displayName = modelItem.DisplayName;
                    //Create a group
                    if (!groups.TryGetValue(modelItem, out currentGroup))
                    {
                        currentGroup = new ClashResultGroup();
                        currentGroup.DisplayName = string.Format("Группа #{0}", (++num).ToString());
                        groups.Add(modelItem, currentGroup);
                    }
                    //Add to the group
                    currentGroup.Children.Add(copiedResult);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("test");
                    ClashResultGroup oneClashResultGroup = new ClashResultGroup();
                    oneClashResultGroup.DisplayName = "Empty clash";
                    oneClashResultGroup.Children.Add(copiedResult);
                    emptyClashResultGroups.Add(oneClashResultGroup);
                }
            }
            foreach (ClashResultGroup g in groups.Values.ToList())
            {
                g.DisplayName = string.Format("{0} ({1})", g.DisplayName, g.Children.Count.ToString());
            }
            List<ClashResultGroup> allGroups = groups.Values.ToList();
            allGroups.AddRange(emptyClashResultGroups);
            return allGroups;
        }
        private static List<ClashResultGroup> GroupByLevel(List<ClashResult> results, string initialName)
        {
            //I already checked if it exists
            GridSystem gridSystem = Application.MainDocument.Grids.ActiveSystem;
            Dictionary<GridLevel, ClashResultGroup> groups = new Dictionary<GridLevel, ClashResultGroup>();
            ClashResultGroup currentGroup;

            //Create a group for the null GridIntersection
            ClashResultGroup nullGridGroup = new ClashResultGroup();
            nullGridGroup.DisplayName = initialName + "No Level";
            int num = 0;
            foreach (ClashResult result in results)
            {
                //Cannot add original result to new clash test, so I create a copy
                ClashResult copiedResult = (ClashResult)result.CreateCopy();

                if (gridSystem.ClosestIntersection(copiedResult.Center) != null)
                {
                    GridLevel closestLevel = gridSystem.ClosestIntersection(copiedResult.Center).Level;

                    if (!groups.TryGetValue(closestLevel, out currentGroup))
                    {
                        currentGroup = new ClashResultGroup();
                        string displayName = closestLevel.DisplayName;
                        currentGroup.DisplayName = string.Format("Группа #{0}", (++num).ToString());
                        /*
                        if (string.IsNullOrEmpty(displayName)) { displayName = "Unnamed Level"; }
                        currentGroup.DisplayName = initialName + displayName;
                        */
                        groups.Add(closestLevel, currentGroup);
                    }
                    currentGroup.Children.Add(copiedResult);
                }
                else
                {
                    nullGridGroup.Children.Add(copiedResult);
                }
            }
            foreach (ClashResultGroup g in groups.Values.ToList())
            {
                g.DisplayName = string.Format("{0} ({1})", g.DisplayName, g.Children.Count.ToString());
            }
            IOrderedEnumerable<KeyValuePair<GridLevel, ClashResultGroup>> list = groups.OrderBy(key => key.Key.Elevation);
            groups = list.ToDictionary((keyItem) => keyItem.Key, (valueItem) => valueItem.Value);

            List<ClashResultGroup> groupsByLevel = groups.Values.ToList();
            if (nullGridGroup.Children.Count != 0) groupsByLevel.Add(nullGridGroup);

            return groupsByLevel;
        }

        private static List<ClashResultGroup> GroupByGridIntersection(List<ClashResult> results, string initialName)
        {
            //I already check if it exists
            GridSystem gridSystem = Application.MainDocument.Grids.ActiveSystem;
            Dictionary<GridIntersection, ClashResultGroup> groups = new Dictionary<GridIntersection, ClashResultGroup>();
            ClashResultGroup currentGroup;

            //Create a group for the null GridIntersection
            ClashResultGroup nullGridGroup = new ClashResultGroup();
            nullGridGroup.DisplayName = initialName + "No Grid intersection";
            int num = 0;
            foreach (ClashResult result in results)
            {
                //Cannot add original result to new clash test, so I create a copy
                ClashResult copiedResult = (ClashResult)result.CreateCopy();

                if (gridSystem.ClosestIntersection(copiedResult.Center) != null)
                {
                    GridIntersection closestGridIntersection = gridSystem.ClosestIntersection(copiedResult.Center);

                    if (!groups.TryGetValue(closestGridIntersection, out currentGroup))
                    {
                        currentGroup = new ClashResultGroup();
                        string displayName = closestGridIntersection.DisplayName;
                        currentGroup.DisplayName = string.Format("Группа #{0}", (++num).ToString());
                        /*
                        if (string.IsNullOrEmpty(displayName)) { displayName = "Unnamed Grid Intersection"; }
                        currentGroup.DisplayName = initialName + displayName;
                        */
                        groups.Add(closestGridIntersection, currentGroup);
                    }
                    currentGroup.Children.Add(copiedResult);
                }
                else
                {
                    nullGridGroup.Children.Add(copiedResult);
                }
            }
            foreach (ClashResultGroup g in groups.Values.ToList())
            {
                g.DisplayName = string.Format("{0} ({1})", g.DisplayName, g.Children.Count.ToString());
            }
            IOrderedEnumerable<KeyValuePair<GridIntersection, ClashResultGroup>> list = groups.OrderBy(key => key.Key.Position.X).OrderBy(key => key.Key.Level.Elevation);
            groups = list.ToDictionary((keyItem) => keyItem.Key, (valueItem) => valueItem.Value);

            List<ClashResultGroup> groupsByGridIntersection = groups.Values.ToList();
            if (nullGridGroup.Children.Count != 0) groupsByGridIntersection.Add(nullGridGroup);

            return groupsByGridIntersection;
        }

        private static List<ClashResultGroup> GroupByElementOfAGivenSelection(List<ClashResult> results, GroupingMode mode, string initialName)
        {
            Dictionary<ModelItem, ClashResultGroup> groups = new Dictionary<ModelItem, ClashResultGroup>();
            ClashResultGroup currentGroup;
            List<ClashResultGroup> emptyClashResultGroups = new List<ClashResultGroup>();
            int num = 0;
            foreach (ClashResult result in results)
            {
                //Cannot add original result to new clash test, so I create a copy
                ClashResult copiedResult = (ClashResult)result.CreateCopy();
                ModelItem modelItem = null;

                if (mode == GroupingMode.SelectionA)
                {
                    if (copiedResult.CompositeItem1 != null)
                    {
                        modelItem = GetSignificantAncestorOrSelf(copiedResult.CompositeItem1);
                    }
                    else if (copiedResult.CompositeItem2 != null)
                    {
                        modelItem = GetSignificantAncestorOrSelf(copiedResult.CompositeItem2);
                    }
                }
                else if (mode == GroupingMode.SelectionB)
                {
                    if (copiedResult.CompositeItem2 != null)
                    {
                        modelItem = GetSignificantAncestorOrSelf(copiedResult.CompositeItem2);
                    }
                    else if (copiedResult.CompositeItem1 != null)
                    {
                        modelItem = GetSignificantAncestorOrSelf(copiedResult.CompositeItem1);
                    }
                }
                string displayName = "Empty clash";
                if (modelItem != null)
                {
                    displayName = modelItem.DisplayName;
                    //Create a group
                    if (!groups.TryGetValue(modelItem, out currentGroup))
                    {
                        currentGroup = new ClashResultGroup();
                        currentGroup.DisplayName = string.Format("Группа #{0}", (++num).ToString());
                        /*
                        if (string.IsNullOrEmpty(displayName)) { displayName = modelItem.Parent.DisplayName; }
                        if (string.IsNullOrEmpty(displayName)) { displayName = "Unnamed Parent"; }
                        currentGroup.DisplayName = initialName + displayName;
                        */
                        groups.Add(modelItem, currentGroup);
                    }
                    //Add to the group
                    currentGroup.Children.Add(copiedResult);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("test");
                    ClashResultGroup oneClashResultGroup = new ClashResultGroup();
                    oneClashResultGroup.DisplayName = "Empty clash";
                    oneClashResultGroup.Children.Add(copiedResult);
                    emptyClashResultGroups.Add(oneClashResultGroup);
                }
            }
            foreach (ClashResultGroup g in groups.Values.ToList())
            {
                g.DisplayName = string.Format("{0} ({1})", g.DisplayName, g.Children.Count.ToString());
            }
            List<ClashResultGroup> allGroups = groups.Values.ToList();
            allGroups.AddRange(emptyClashResultGroups);
            return allGroups;
        }

        private static List<ClashResultGroup> GroupByProperties(List<ClashResult> results, GroupingMode mode, string initialName)
        {
            Dictionary<string, ClashResultGroup> groups = new Dictionary<string, ClashResultGroup>();
            ClashResultGroup currentGroup;
            int num = 0;
            foreach (ClashResult result in results)
            {
                //Cannot add original result to new clash test, so I create a copy
                ClashResult copiedResult = (ClashResult)result.CreateCopy();
                string clashProperty = null;

                if (mode == GroupingMode.ApprovedBy)
                {
                    clashProperty = copiedResult.ApprovedBy;
                }
                else if (mode == GroupingMode.AssignedTo)
                {
                    clashProperty = copiedResult.AssignedTo;
                }
                else if (mode == GroupingMode.Status)
                {
                    clashProperty = copiedResult.Status.ToString();
                }

                if (string.IsNullOrEmpty(clashProperty)) { clashProperty = "Unspecified"; }

                if (!groups.TryGetValue(clashProperty, out currentGroup))
                {
                    currentGroup = new ClashResultGroup();
                    currentGroup.DisplayName = string.Format("Группа #{0}", (++num).ToString());
                    /*
                    currentGroup.DisplayName = initialName + clashProperty;
                    */
                    groups.Add(clashProperty, currentGroup);
                }
                currentGroup.Children.Add(copiedResult);
            }
            foreach (ClashResultGroup g in groups.Values.ToList())
            {
                g.DisplayName = string.Format("{0} ({1})", g.DisplayName, g.Children.Count.ToString());
            }
            return groups.Values.ToList();
        }

        #endregion


        #region helpers






        private static void ProcessClashGroup(List<ClashResultGroup> clashGroups, List<ClashResult> ungroupedClashResults, ClashTest selectedClashTest)
        {
            using (Transaction tx = Application.MainDocument.BeginTransaction("Group clashes"))
            {
                ClashTest copiedClashTest = (ClashTest)selectedClashTest.CreateCopyWithoutChildren();
                //When we replace theTest with our new test, theTest will be disposed. If the operation is cancelled, we need a non-disposed copy of theTest with children to sub back in.
                ClashTest BackupTest = (ClashTest)selectedClashTest.CreateCopy();
                DocumentClash documentClash = Application.MainDocument.GetClash();
                int indexOfClashTest = documentClash.TestsData.Tests.IndexOf(selectedClashTest);
                documentClash.TestsData.TestsReplaceWithCopy(indexOfClashTest, copiedClashTest);

                int CurrentProgress = 0;
                int TotalProgress = ungroupedClashResults.Count + clashGroups.Count;
                Progress ProgressBar = Application.BeginProgress("Copying Results", "Copying results from " + selectedClashTest.DisplayName + " to the Group Clashes pane...");
                foreach (ClashResultGroup clashResultGroup in clashGroups)
                {
                    if (ProgressBar.IsCanceled) break;
                    documentClash.TestsData.TestsAddCopy((GroupItem)documentClash.TestsData.Tests[indexOfClashTest], clashResultGroup);
                    CurrentProgress++;
                    ProgressBar.Update((double)CurrentProgress / TotalProgress);
                }
                foreach (ClashResult clashResult in ungroupedClashResults)
                {
                    if (ProgressBar.IsCanceled) break;
                    documentClash.TestsData.TestsAddCopy((GroupItem)documentClash.TestsData.Tests[indexOfClashTest], clashResult);
                    CurrentProgress++;
                    ProgressBar.Update((double)CurrentProgress / TotalProgress);
                }
                if (ProgressBar.IsCanceled) documentClash.TestsData.TestsReplaceWithCopy(indexOfClashTest, BackupTest);
                tx.Commit();
                Application.EndProgress();
            }
        }








        private static void AddUngroupedClashes(List<ClashResult> unfilteredClashResults, ClashTest selectedClashTest)
{
    using (Transaction tx = Application.MainDocument.BeginTransaction("Add unfiltered clashes"))
    {
        DocumentClash documentClash = Application.MainDocument.GetClash();
        int indexOfClashTest = documentClash.TestsData.Tests.IndexOf(selectedClashTest);
        GroupItem targetGroup = documentClash.TestsData.Tests[indexOfClashTest] as GroupItem;

        if (targetGroup == null)
        {
            throw new InvalidOperationException("GroupItem not found for the selected ClashTest.");
        }

        foreach (ClashResult clashResult in unfilteredClashResults)
        {
            documentClash.TestsData.TestsAddCopy(targetGroup, clashResult);
        }

        tx.Commit();
    }
}


        private static List<ClashResult> RemoveOneClashGroup(ref List<ClashResultGroup> clashResultGroups)
        {
            List<ClashResult> ungroupedClashResults = new List<ClashResult>();
            List<ClashResultGroup> temporaryClashResultGroups = new List<ClashResultGroup>();
            temporaryClashResultGroups.AddRange(clashResultGroups);

            foreach (ClashResultGroup group in temporaryClashResultGroups)
            {
                if (group.Children.Count == 1)
                {
                    ClashResult result = (ClashResult)group.Children.FirstOrDefault();
                    ungroupedClashResults.Add(result);
                    clashResultGroups.Remove(group);
                }
            }

            return ungroupedClashResults;
        }

        private static IEnumerable<ClashResult> GetIndividualClashResults(ClashTest clashTest, bool keepExistingGroup)
        {
            for (var i = 0; i < clashTest.Children.Count; i++)
            {
                if (clashTest.Children[i].IsGroup)
                {
                    if (!keepExistingGroup)
                    {
                        IEnumerable<ClashResult> GroupResults = GetGroupResults((ClashResultGroup)clashTest.Children[i]);
                        foreach (ClashResult clashResult in GroupResults)
                        {
                            yield return clashResult;
                        }
                    }
                }
                else yield return (ClashResult)clashTest.Children[i];
            }
        }

        private static IEnumerable<ClashResultGroup> BackupExistingClashGroups(ClashTest clashTest)
        {
            for (var i = 0; i < clashTest.Children.Count; i++)
            {
                if (clashTest.Children[i].IsGroup)
                {
                    yield return (ClashResultGroup)clashTest.Children[i].CreateCopy();
                }
            }
        }

        private static IEnumerable<ClashResult> GetGroupResults(ClashResultGroup clashResultGroup)
        {
            for (var i = 0; i < clashResultGroup.Children.Count; i++)
            {
                yield return (ClashResult)clashResultGroup.Children[i];
            }
        }

        private static ModelItem GetSignificantAncestorOrSelf(ModelItem item)
        {
            try
            {
                ModelItem originalItem = item;
                ModelItem currentComposite = null;

                //Get last composite item.
                while (item.Parent != null)
                {
                    item = item.Parent;
                    if (item.IsComposite) currentComposite = item;
                }

                return currentComposite ?? originalItem;
            }
            catch (Exception)
            {
                return null;
            }
        }
        #endregion

    }

}
