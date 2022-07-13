using System;

namespace Navisworks.Clash.Exporter.Extensions.Attributes
{
    public class TableNameAttribute : Attribute
    {
        public string Name { get; set; }

        public TableNameAttribute(string name)
        {
            Name = name;
        }
    }
}