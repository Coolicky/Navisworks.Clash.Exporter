using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api.Clash;
using Navisworks.Clash.Exporter.Extensions;
using Navisworks.Clash.Exporter.Extensions.Attributes;

namespace Navisworks.Clash.Exporter.Data
{
    [TableName("Clash Groups")]
    public class ClashGroupDto
    {
        public ClashGroupDto(ClashResultGroup group, Guid savedItemGuid)
        {
            TestGuid = savedItemGuid.ToString();
            DisplayName = group.DisplayName;
            Guid = group.Guid.ToString();
            ClashesCount = group.Children?.Count;
            ApprovedBy = group.ApprovedBy;
            ApprovedTime = group.ApprovedTime;
            AssignedTo = group.AssignedTo;
            try
            {
                Center = $"{group.Center.X}, {group.Center.Y}, {group.Center.Z}";
            }
            catch
            {
                // ignored
            }

            CreatedTime = group.CreatedTime;
            Description = group.Description;
            Distance = group.Distance.ToCurrentUnits();
            Status = group.Status.ToString();
            Comments = group.Comments?.Select(r => new CommentDto(r, Guid)).ToList();
        }

        [ColumnName("Test Guid")] public string TestGuid { get; }
        [ColumnName("Name")] public string DisplayName { get; }
        [ColumnName("Guid")] public string Guid { get; }
        [ColumnName("No. Clashes")] public int? ClashesCount { get; }
        [ColumnName("Approved By")] public string ApprovedBy { get; }
        [ColumnName("Approved Time")] public DateTime? ApprovedTime { get; }
        [ColumnName("Assigned To")] public string AssignedTo { get; }
        [ColumnName("Center")] public string Center { get; }
        [ColumnName("Created Time")] public DateTime? CreatedTime { get; }
        [ColumnName("Description")] public string Description { get; }
        [ColumnName("Distance")] public double Distance { get; }
        [ColumnName("Status")] public string Status { get; }
        [IgnoreColumn] public List<CommentDto> Comments { get; }
    }
}