using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;

namespace Navisworks.Clash.Exporter.ClashGrouping
{
    public static class GroupingFunctions
    {
        public static void GroupClashes(ClashTest selectedClashTest, GroupingMode groupingMode,
            GroupingMode subgroupMode, bool keepExistingGroups)
        {
            //Get existing clash result
            var clashResults = GetIndividualClashResults(selectedClashTest, keepExistingGroups).ToList();
            var clashResultGroups = new List<ClashResultGroup>();

            //Create groups according to the first grouping mode
            CreateGroup(ref clashResultGroups, groupingMode, clashResults, "");

            //Optionally, create subgroups
            if (subgroupMode != GroupingMode.None)
            {
                CreateSubGroups(ref clashResultGroups, subgroupMode);
            }

            //Remove groups with only one clash
            // var ungroupedClashResults = RemoveOneClashGroup(ref clashResultGroups);

            //Backup the existing group, if necessary
            if (keepExistingGroups) clashResultGroups.AddRange(BackupExistingClashGroups(selectedClashTest));

            //Process these groups and clashes into the clash test
            // ProcessClashGroup(clashResultGroups, ungroupedClashResults, selectedClashTest);
            ProcessClashGroup(clashResultGroups, new List<ClashResult>(), selectedClashTest);
        }

        private static void CreateGroup(ref List<ClashResultGroup> clashResultGroups, GroupingMode groupingMode,
            List<ClashResult> clashResults, string initialName)
        {
            //group all clashes
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
                case GroupingMode.ModelA:
                case GroupingMode.ModelB:
                    clashResultGroups = GroupByElementOfAGivenModel(clashResults, groupingMode, initialName);
                    break;
                case GroupingMode.ApprovedBy:
                case GroupingMode.AssignedTo:
                case GroupingMode.Status:
                    clashResultGroups = GroupByProperties(clashResults, groupingMode, initialName);
                    break;
                case GroupingMode.ItemTypeA:
                    clashResultGroups = GroupByClass(clashResults, groupingMode, initialName);
                    break;
                case GroupingMode.ItemTypeB:
                    clashResultGroups = GroupByClass(clashResults, groupingMode, initialName);
                    break;
                default:
                    return;
            }
        }

        // ReSharper disable once CognitiveComplexity
        private static List<ClashResultGroup> GroupByClass(IEnumerable<ClashResult> results, GroupingMode mode, string initialName)
        {
            var groups = new Dictionary<string, ClashResultGroup>();
            var emptyClashResultGroups = new List<ClashResultGroup>();

            foreach (var result in results)
            {
                //Cannot add original result to new clash test, so I create a copy
                var copiedResult = (ClashResult)result.CreateCopy();
                // ModelItem modelItem = null;
                ModelItem rootModel = null;

                if (mode == GroupingMode.ItemTypeA)
                {
                    rootModel = GetGuidAncestor(result.Item1);
                }
                else if (mode == GroupingMode.ItemTypeB)
                {
                    rootModel = GetGuidAncestor(result.Item2);
                }

                var className = "Unknown";

                if (rootModel != null)
                {
                    className = rootModel.ClassDisplayName;
                    className = className.Split(':').FirstOrDefault();
                    if (string.IsNullOrEmpty(className)) className = "Unknown";
                }
                
                //Create or find a group
                if (!groups.TryGetValue(className, out var currentGroup))
                {
                    currentGroup = new ClashResultGroup();

                    currentGroup.DisplayName = initialName + className;
                    groups.Add(className, currentGroup);
                }

                //Add to the group
                currentGroup.Children.Add(copiedResult);
            }

            var allGroups = groups.Values.ToList();
            allGroups.AddRange(emptyClashResultGroups);
            return allGroups;
        }

        private static void CreateSubGroups(ref List<ClashResultGroup> clashResultGroups, GroupingMode mode)
        {
            var clashResultSubGroups = new List<ClashResultGroup>();

            foreach (var group in clashResultGroups)
            {
                var clashResults = group.Children.OfType<ClashResult>().ToList();

                var clashResultTempSubGroups = new List<ClashResultGroup>();
                CreateGroup(ref clashResultTempSubGroups, mode, clashResults, group.DisplayName + "_");
                clashResultSubGroups.AddRange(clashResultTempSubGroups);
            }

            clashResultGroups = clashResultSubGroups;
        }
        

        public static void UnGroupClashes(ClashTest selectedClashTest)
        {
            var groups = new List<ClashResultGroup>();
            var results = GetIndividualClashResults(selectedClashTest,false).ToList();
            var copiedResult = results.Select(result => (ClashResult)result.CreateCopy()).ToList();

            //Process this empty group list and clashes into the clash test
            ProcessClashGroup(groups, copiedResult, selectedClashTest);
        }

        #region grouping functions

        // ReSharper disable once CognitiveComplexity
        private static List<ClashResultGroup> GroupByLevel(IEnumerable<ClashResult> results, string initialName)
        {
            //I already checked if it exists
            var gridSystem = Application.MainDocument.Grids.ActiveSystem;
            var groups = new Dictionary<GridLevel, ClashResultGroup>();

            //Create a group for the null GridIntersection
            var nullGridGroup = new ClashResultGroup();
            nullGridGroup.DisplayName = initialName + "No Level";

            foreach (var copiedResult in results.Select(result => (ClashResult)result.CreateCopy()))
            {
                if (gridSystem.ClosestIntersection(copiedResult.Center) != null)
                {
                    var closestLevel = gridSystem.ClosestIntersection(copiedResult.Center).Level;

                    if (!groups.TryGetValue(closestLevel, out var currentGroup))
                    {
                        currentGroup = new ClashResultGroup();
                        var displayName = closestLevel.DisplayName;
                        if (string.IsNullOrEmpty(displayName))
                        {
                            displayName = "Unnamed Level";
                        }

                        currentGroup.DisplayName = initialName + displayName;
                        groups.Add(closestLevel, currentGroup);
                    }

                    currentGroup.Children.Add(copiedResult);
                }
                else
                {
                    nullGridGroup.Children.Add(copiedResult);
                }
            }

            IOrderedEnumerable<KeyValuePair<GridLevel, ClashResultGroup>> list =
                groups.OrderBy(key => key.Key.Elevation);
            groups = list.ToDictionary((keyItem) => keyItem.Key, (valueItem) => valueItem.Value);

            List<ClashResultGroup> groupsByLevel = groups.Values.ToList();
            if (nullGridGroup.Children.Count != 0) groupsByLevel.Add(nullGridGroup);

            return groupsByLevel;
        }

        // ReSharper disable once CognitiveComplexity
        private static List<ClashResultGroup> GroupByGridIntersection(List<ClashResult> results, string initialName)
        {
            //I already check if it exists
            var gridSystem = Application.MainDocument.Grids.ActiveSystem;
            var groups = new Dictionary<GridIntersection, ClashResultGroup>();

            //Create a group for the null GridIntersection
            var nullGridGroup = new ClashResultGroup();
            nullGridGroup.DisplayName = initialName + "No Grid intersection";

            foreach (var result in results)
            {
                //Cannot add original result to new clash test, so I create a copy
                var copiedResult = (ClashResult)result.CreateCopy();

                if (gridSystem.ClosestIntersection(copiedResult.Center) != null)
                {
                    GridIntersection closestGridIntersection = gridSystem.ClosestIntersection(copiedResult.Center);

                    if (!groups.TryGetValue(closestGridIntersection, out var currentGroup))
                    {
                        currentGroup = new ClashResultGroup();
                        string displayName = closestGridIntersection.DisplayName;
                        if (string.IsNullOrEmpty(displayName))
                        {
                            displayName = "Unnamed Grid Intersection";
                        }

                        currentGroup.DisplayName = initialName + displayName;
                        groups.Add(closestGridIntersection, currentGroup);
                    }

                    currentGroup.Children.Add(copiedResult);
                }
                else
                {
                    nullGridGroup.Children.Add(copiedResult);
                }
            }

            var list = groups.OrderBy(key => key.Key.Position.X).ThenBy(key => key.Key.Level.Elevation);
            groups = list.ToDictionary((keyItem) => keyItem.Key, (valueItem) => valueItem.Value);

            var groupsByGridIntersection = groups.Values.ToList();
            if (nullGridGroup.Children.Count != 0) groupsByGridIntersection.Add(nullGridGroup);

            return groupsByGridIntersection;
        }

        // ReSharper disable once CognitiveComplexity
        private static List<ClashResultGroup> GroupByElementOfAGivenSelection(List<ClashResult> results,
            GroupingMode mode, string initialName)
        {
            var groups = new Dictionary<ModelItem, ClashResultGroup>();
            var emptyClashResultGroups = new List<ClashResultGroup>();

            foreach (var result in results)
            {
                //Cannot add original result to new clash test, so I create a copy
                var copiedResult = (ClashResult)result.CreateCopy();
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

                if (modelItem != null)
                {
                    var displayName = modelItem.DisplayName;
                    //Create a group
                    if (!groups.TryGetValue(modelItem, out var currentGroup))
                    {
                        currentGroup = new ClashResultGroup();
                        if (string.IsNullOrEmpty(displayName))
                        {
                            displayName = modelItem.Parent.DisplayName;
                        }

                        if (string.IsNullOrEmpty(displayName))
                        {
                            displayName = "Unnamed Parent";
                        }

                        currentGroup.DisplayName = initialName + displayName;
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

            List<ClashResultGroup> allGroups = groups.Values.ToList();
            allGroups.AddRange(emptyClashResultGroups);
            return allGroups;
        }

        // ReSharper disable once CognitiveComplexity
        private static List<ClashResultGroup> GroupByElementOfAGivenModel(List<ClashResult> results, GroupingMode mode,
            string initialName)
        {
            var groups = new Dictionary<ModelItem, ClashResultGroup>();
            var emptyClashResultGroups = new List<ClashResultGroup>();

            foreach (var result in results)
            {
                //Cannot add original result to new clash test, so I create a copy
                var copiedResult = (ClashResult)result.CreateCopy();
                // ModelItem modelItem = null;
                ModelItem rootModel = null;

                if (mode == GroupingMode.ModelA)
                {
                    if (copiedResult.CompositeItem1 != null)
                    {
                        rootModel = GetFileAncestor(copiedResult.CompositeItem1);
                    }
                    else if (copiedResult.CompositeItem2 != null)
                    {
                        rootModel = GetFileAncestor(copiedResult.CompositeItem2);
                    }
                }
                else if (mode == GroupingMode.ModelB)
                {
                    if (copiedResult.CompositeItem2 != null)
                    {
                        rootModel = GetFileAncestor(copiedResult.CompositeItem2);
                    }
                    else if (copiedResult.CompositeItem1 != null)
                    {
                        rootModel = GetFileAncestor(copiedResult.CompositeItem1);
                    }
                }

                if (rootModel != null)
                {
                    var displayName = rootModel.DisplayName;
                    //Create a group
                    if (!groups.TryGetValue(rootModel, out var currentGroup))
                    {
                        currentGroup = new ClashResultGroup();
                        // if (string.IsNullOrEmpty(displayName)) { displayName = rootModel.Parent.DisplayName; }
                        if (string.IsNullOrEmpty(displayName))
                        {
                            displayName = "Unnamed Model";
                        }

                        currentGroup.DisplayName = initialName + displayName;
                        groups.Add(rootModel, currentGroup);
                    }

                    //Add to the group
                    currentGroup.Children.Add(copiedResult);
                }
                else
                {
                    var oneClashResultGroup = new ClashResultGroup();
                    oneClashResultGroup.DisplayName = "Empty clash";
                    oneClashResultGroup.Children.Add(copiedResult);
                    emptyClashResultGroups.Add(oneClashResultGroup);
                }
            }

            var allGroups = groups.Values.ToList();
            allGroups.AddRange(emptyClashResultGroups);
            return allGroups;
        }

        private static List<ClashResultGroup> GroupByProperties(List<ClashResult> results, GroupingMode mode,
            string initialName)
        {
            var groups = new Dictionary<string, ClashResultGroup>();

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

                if (string.IsNullOrEmpty(clashProperty))
                {
                    clashProperty = "Unspecified";
                }

                if (!groups.TryGetValue(clashProperty, out var currentGroup))
                {
                    currentGroup = new ClashResultGroup();
                    currentGroup.DisplayName = initialName + clashProperty;
                    groups.Add(clashProperty, currentGroup);
                }

                currentGroup.Children.Add(copiedResult);
            }

            return groups.Values.ToList();
        }

        #endregion


        #region helpers

        private static void ProcessClashGroup(IReadOnlyCollection<ClashResultGroup> clashGroups,
            IReadOnlyCollection<ClashResult> ungroupedClashResults, ClashTest selectedClashTest)
        {
            using (var tx = Application.MainDocument.BeginTransaction("Group clashes"))
            {
                var copiedClashTest = (ClashTest)selectedClashTest.CreateCopyWithoutChildren();
                //When we replace theTest with our new test, theTest will be disposed. If the operation is cancelled, we need a non-disposed copy of theTest with children to sub back in.
                var backupTest = (ClashTest)selectedClashTest.CreateCopy();
                var documentClash = Application.MainDocument.GetClash();
                var indexOfClashTest = documentClash.TestsData.Tests.IndexOf(selectedClashTest);
                documentClash.TestsData.TestsReplaceWithCopy(indexOfClashTest, copiedClashTest);

                var currentProgress = 0;
                var totalProgress = ungroupedClashResults.Count + clashGroups.Count;
                var progressBar = Application.BeginProgress("Copying Results",
                    "Copying results from " + selectedClashTest.DisplayName + " to the Group Clashes pane...");
                foreach (var clashResultGroup in clashGroups.TakeWhile(clashResultGroup => !progressBar.IsCanceled))
                {
                    documentClash.TestsData.TestsAddCopy((GroupItem)documentClash.TestsData.Tests[indexOfClashTest],
                        clashResultGroup);
                    currentProgress++;
                    progressBar.Update((double)currentProgress / totalProgress);
                }

                foreach (var clashResult in ungroupedClashResults.TakeWhile(clashResult => !progressBar.IsCanceled))
                {
                    documentClash.TestsData.TestsAddCopy((GroupItem)documentClash.TestsData.Tests[indexOfClashTest],
                        clashResult);
                    currentProgress++;
                    progressBar.Update((double)currentProgress / totalProgress);
                }

                if (progressBar.IsCanceled) documentClash.TestsData.TestsReplaceWithCopy(indexOfClashTest, backupTest);
                tx.Commit();
                Application.EndProgress();
            }
        }

        private static IEnumerable<ClashResult> GetIndividualClashResults(ClashTest clashTest, bool keepExistingGroup)
        {
            foreach (var savedItem in clashTest.Children)
            {
                if (savedItem.IsGroup)
                {
                    if (keepExistingGroup) continue;
                    var groupResults = GetGroupResults((ClashResultGroup)savedItem);
                    foreach (var clashResult in groupResults)
                    {
                        yield return clashResult;
                    }
                }
                else yield return (ClashResult)savedItem;
            }
        }

        private static IEnumerable<ClashResultGroup> BackupExistingClashGroups(ClashTest clashTest)
        {
            foreach (var savedItem in clashTest.Children)
            {
                if (!savedItem.IsGroup) continue;
                yield return (ClashResultGroup)savedItem.CreateCopy();
            }
        }

        private static IEnumerable<ClashResult> GetGroupResults(ClashResultGroup clashResultGroup)
        {
            return clashResultGroup.Children.Cast<ClashResult>();
        }

        private static ModelItem GetSignificantAncestorOrSelf(ModelItem item)
        {
            var originalItem = item;
            ModelItem currentComposite = null;

            //Get last composite item.
            while (item.Parent != null)
            {
                item = item.Parent;
                if (item.IsComposite) currentComposite = item;
            }

            return currentComposite ?? originalItem;
        }

        private static ModelItem GetFileAncestor(ModelItem item)
        {
            var originalItem = item;

            ModelItem currentComposite = null;

            //Get last composite item.
            while (item.Parent != null)
            {
                item = item.Parent;
                if (!item.HasModel) continue;
                currentComposite = item;
                break;
            }

            return currentComposite ?? originalItem;
        }
        

        private static ModelItem GetGuidAncestor(ModelItem modelItem)
        {
            if (modelItem == null) return null;
            while (modelItem.InstanceGuid == new Guid("00000000-0000-0000-0000-000000000000"))
            {
                modelItem = modelItem.Parent;
            }

            return modelItem;
        }

        #endregion
    }
}