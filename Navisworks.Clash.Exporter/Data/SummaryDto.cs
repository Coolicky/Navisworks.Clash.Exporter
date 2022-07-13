using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api.Clash;
using Navisworks.Clash.Exporter.Extensions.Attributes;

namespace Navisworks.Clash.Exporter.Data
{
    [TableName("Test Summary")]
    public class SummaryDto
    {
        public SummaryDto(ClashTest test)
        {
            TestName = test.DisplayName;
            New = GetCount(test, ClashResultStatus.New);
            Active = GetCount(test, ClashResultStatus.Active);
            Reviewed = GetCount(test, ClashResultStatus.Reviewed);
            Approved = GetCount(test, ClashResultStatus.Approved);
            Resolved = GetCount(test, ClashResultStatus.Resolved);
            Total = New + Active + Reviewed + Approved + Resolved;
        }

        [ColumnName("Test Name")] public string TestName { get; }
        [ColumnName("New")] public int New { get; }
        [ColumnName("Active")] public int Active { get; }
        [ColumnName("Reviewed")] public int Reviewed { get; }
        [ColumnName("Approved")] public int Approved { get; }
        [ColumnName("Resolved")] public int Resolved { get; }
        [ColumnName("Total")] public int Total { get; }

        private int GetCount(ClashTest test, ClashResultStatus status)
        {
            var children = new List<ClashResult>();

            try
            {
                var groups = test.Children
                    .Where(r => r.GetType() == typeof(ClashResultGroup))
                    .Cast<ClashResultGroup>()
                    .ToList();
                children.AddRange(groups
                    .SelectMany(group => group.Children
                        .Where(r => r.GetType() == typeof(ClashResult))
                        .Cast<ClashResult>()));

                children.AddRange(test.Children.Where(r => r.GetType() == typeof(ClashResult)).Cast<ClashResult>());
            }
            catch
            {
                // ignored
            }

            return children.Count(r => r.Status == status);
        }
    }
}