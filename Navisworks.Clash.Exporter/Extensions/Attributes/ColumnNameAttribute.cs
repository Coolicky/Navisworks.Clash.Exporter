using System;

namespace Navisworks.Clash.Exporter.Extensions.Attributes
{
    public class ColumnNameAttribute : Attribute
    {
        public string Name { get; set; }

        public ColumnNameAttribute(string name)
        {
            Name = name;
        }
    }
}