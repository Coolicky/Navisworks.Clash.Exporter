using System.ComponentModel;

namespace Navisworks.Clash.Exporter.ClashGrouping
{
    public enum GroupingMode
    {
        [Description("<None>")] None,
        [Description("Level")] Level,
        [Description("Grid Intersection")] GridIntersection,
        [Description("Selection A")] SelectionA,
        [Description("Selection B")] SelectionB,
        [Description("Model A")] ModelA,
        [Description("Model B")] ModelB,
        [Description("Assigned To")] AssignedTo,
        [Description("Approved By")] ApprovedBy,
        [Description("Status")] Status,
        [Description("Item Type A")] ItemTypeA,
        [Description("Item Type B")] ItemTypeB
    }
}