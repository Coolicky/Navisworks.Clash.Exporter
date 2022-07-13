using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using Navisworks.Clash.Exporter.Extensions;
using Navisworks.Clash.Exporter.Extensions.Attributes;

namespace Navisworks.Clash.Exporter.Data
{
    [TableName("Clash Tests")]
    public class ClashTestDto
    {
        public ClashTestDto(ClashTest test)
        {
            DisplayName = test.DisplayName;
            Guid = test.Guid.ToString();
            ClashesCount = GetAllClashesCount(test.Children);
            Status = test.Status.ToString();
            LastRun = test.LastRun;
            TestType = test.TestType.ToString();
            Tolerance = test.Tolerance.ToCurrentUnits();
            Comments = test.Comments.Select(r => new CommentDto(r, Guid)).ToList();
        }

        private int GetAllClashesCount(SavedItemCollection children)
        {
            var count = 0;

            try
            {
                var groups = children
                    .Where(r => r.GetType() == typeof(ClashResultGroup))
                    .Cast<ClashResultGroup>()
                    .ToList();
                count += groups.Sum(group => group.Children.Count);
            }
            catch
            {
                // ignored
            }

            try
            {
                var clashes = children
                    .Where(r => r.GetType() == typeof(ClashResult))
                    .ToList();
                count += clashes.Count;
            }
            catch
            {
                // ignored
            }

            return count;
        }

        [ColumnName("Name")] public string DisplayName { get; }
        [ColumnName("Guid")] public string Guid { get; }
        [ColumnName("No. Clashes")] public int ClashesCount { get; }
        [ColumnName("Status")] public string Status { get; }
        [ColumnName("Last Run")] public DateTime? LastRun { get; }
        [ColumnName("Test Type")] public string TestType { get; }
        [ColumnName("Tolerance")] public double Tolerance { get; }
        [IgnoreColumn] public List<CommentDto> Comments { get; }
    }
}