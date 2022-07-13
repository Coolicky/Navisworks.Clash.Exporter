using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using Navisworks.Clash.Exporter.Extensions;
using Navisworks.Clash.Exporter.Extensions.Attributes;

namespace Navisworks.Clash.Exporter.Data
{
    [TableName("Clash Results")]
    public class ClashResultDto
    {
        public ClashResultDto(ClashResult clash, Guid savedItemGuid)
        {
            TestGuid = savedItemGuid.ToString();
            if (clash.Parent?.GetType() == typeof(ClashResultGroup))
                GroupGuid = clash.Parent?.Guid.ToString();
            DisplayName = clash.DisplayName;
            Guid = clash.Guid.ToString();
            ApprovedBy = clash.ApprovedBy;
            ApprovedTime = clash.ApprovedTime;
            AssignedTo = clash.AssignedTo;
            try
            {
                Center = $"{clash.Center.X}, {clash.Center.Y}, {clash.Center.Z}";
            }
            catch
            {
                // ignored
            }

            CreatedTime = clash.CreatedTime;
            Description = clash.Description;
            Distance = clash.Distance.ToCurrentUnits();
            Status = clash.Status.ToString();
            Comments = clash.Comments?.Select(r => new CommentDto(r, Guid)).ToList();
            Items = new List<ElementDto>();
            if (clash.Item1 != null)
            {
                var item = clash.Item1;
                while (item.InstanceGuid == new Guid("00000000-0000-0000-0000-000000000000"))
                {
                    item = item.Parent;
                }

                Item1Guid = item.InstanceGuid.ToString();
                Items.Add(new ElementDto(item));
            }

            if (clash.Item2 != null)
            {
                var item = clash.Item2;
                while (item.InstanceGuid == new Guid("00000000-0000-0000-0000-000000000000"))
                {
                    item = item.Parent;
                }

                Item2Guid = item.InstanceGuid.ToString();
                Items.Add(new ElementDto(item));
            }

            var gridSystem = Application.MainDocument.Grids.ActiveSystem;
            Level = gridSystem.ClosestIntersection(clash.Center)?.Level?.DisplayName;
            GridIntersection = gridSystem.ClosestIntersection(clash.Center)?.DisplayName
                .Replace($" : {Level}", string.Empty);
        }

        [ColumnName("Test Guid")] public string TestGuid { get; }
        [ColumnName("Group Guid")] public string GroupGuid { get; }
        [ColumnName("Name")] public string DisplayName { get; }
        [ColumnName("Guid")] public string Guid { get; }
        [ColumnName("Approved By")] public string ApprovedBy { get; }
        [ColumnName("Approved Time")] public DateTime? ApprovedTime { get; }
        [ColumnName("Assigned To")] public string AssignedTo { get; }
        [ColumnName("Center")] public string Center { get; }
        [ColumnName("Created Time")] public DateTime? CreatedTime { get; }
        [ColumnName("Description")] public string Description { get; }
        [ColumnName("Distance")] public double Distance { get; }
        [ColumnName("Status")] public string Status { get; }
        [ColumnName("Item 2 Guid")] public string Item1Guid { get; }
        [ColumnName("Item 1 Guid")] public string Item2Guid { get; }
        [ColumnName("Grid Intersection")] public string GridIntersection { get; }
        [ColumnName("Level")] public string Level { get; }
        [IgnoreColumn] public List<ElementDto> Items { get; }
        [IgnoreColumn] public List<CommentDto> Comments { get; }
    }
}