using Autodesk.Navisworks.Api;
using System;

namespace Navisworks.Clash.Exporter.Extensions
{
    public static class NavisworksExtensions
    {
        public static bool Imperial { get; set; }
        public static double ToCurrentUnits(this double value)
        {
            var scaleFactor = UnitConversion.ScaleFactor(Units.Feet, Units.Millimeters);
            if (Imperial)
                scaleFactor = 1;
            return value * scaleFactor;
        }

        public static ModelItem GetUniquelyIdentifiableItem(this ModelItem item, out string uniqueId)
        {
            string id;
            var currentItem = item;

            while (true)
            {
                var guid = currentItem.InstanceGuid; //Revit GUID
                if (guid == Guid.Empty)
                {
                    id = currentItem.FromAutoCAD() ?? currentItem.FromGuidTab() ?? currentItem.FromMicrostation();
                }
                else
                {
                    id = guid.ToString();
                }
                if (id != null) break; // Found a unique identifier
                if (currentItem.Parent == null) // No parent item, stop searching
                {
                    currentItem = null;
                    break;
                }
                currentItem = currentItem.Parent; // Move to parent item
            }

            uniqueId = id;
            return currentItem;
        }

        private static string FromAutoCAD(this ModelItem item)
        {
            var cat =
                item.PropertyCategories.FindCategoryByName("LcOpDwgEntityAttrib") ??
                item.PropertyCategories.FindCategoryByDisplayName("Entity Handle");

            if (cat == null) return null;

            var value = cat.Properties.FindPropertyByName("LcOaNat64AttributeValue") ??
                        cat.Properties.FindPropertyByDisplayName("Value");

            if (value == null) return null;
            return value.Value.ToDisplayString();
        }

        private static string FromGuidTab(this ModelItem item)
        {
            var cat =
                item.PropertyCategories.FindCategoryByName("LcArGUID") ??
                item.PropertyCategories.FindCategoryByDisplayName("GUID");

            if (cat == null) return null;

            var value = cat.Properties.FindPropertyByName("LcOaNat64AttributeValue") ??
                        cat.Properties.FindPropertyByDisplayName("Value");

            if (value == null) return null;
            return value.Value.ToDisplayString();
        }

        private static string FromMicrostation(this ModelItem item)
        {
            var cat = item.PropertyCategories.FindCategoryByDisplayName("Element ID");

            if (cat == null) return null;

            var value = cat.Properties.FindPropertyByName("LcOaNat64AttributeValue") ??
                        cat.Properties.FindPropertyByDisplayName("Value");

            if (value == null) return null;
            return value.Value.ToDisplayString();
        }

        public static T Clamp<T>(this T value, T min, T max) where T : IComparable<T>
        {
            if (value.CompareTo(min) < 0) return min;
            if (value.CompareTo(max) > 0) return max;
            return value;
        }
    }
}