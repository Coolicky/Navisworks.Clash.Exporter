using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.Interop.ComApi;
using Navisworks.Clash.Exporter.Extensions.Attributes;

namespace Navisworks.Clash.Exporter.Data
{
    [TableName("Elements")]
    public class ElementDto
    {
        public ElementDto(ModelItem modelItem)
        {
            while (modelItem.InstanceGuid == new Guid("00000000-0000-0000-0000-000000000000"))
            {
                modelItem = modelItem.Parent;
            }

            Name = modelItem.DisplayName;
            Guid = modelItem.InstanceGuid.ToString();
            ClassName = modelItem.ClassDisplayName;
            SourceFile = GetSourceFile(modelItem);
            QuickProperties = GetQuickProperties(modelItem);
        }

        private string GetSourceFile(ModelItem modelItem)
        {
            try
            {
                return modelItem.Model.SourceFileName;
            }
            catch
            {
                return "Error_NotFound";
            }
        }

        private Dictionary<string, string> GetQuickProperties(ModelItem modelItem)
        {
            var props = new Dictionary<string, string>();

            try
            {
                InwOpState4 oState = ComApiBridge.State;
                var oPath = ComApiBridge.ToInwOaPath(modelItem);
                var quickPropertiesText = oState.SmartTagText(oPath);
                var lines = quickPropertiesText.Split('\n');
                foreach (var line in lines)
                {
                    var split = line.IndexOf(": ", StringComparison.Ordinal);
                    if (split <= 0) continue;

                    var key = line.Substring(0, split);
                    var value = line.Substring(split + 2);
                    if (!props.Keys.Contains(key))
                        props.Add(key, value);
                }
            }
            catch
            {
                // ignored
            }

            return props;
        }

        [ColumnName("Name")] public string Name { get; }
        [ColumnName("Guid")] public string Guid { get; }
        [ColumnName("Class")] public string ClassName { get; }
        [ColumnName("Source File")] public string SourceFile { get; }
        [IgnoreColumn] public Dictionary<string, string> QuickProperties { get; }
    }
}