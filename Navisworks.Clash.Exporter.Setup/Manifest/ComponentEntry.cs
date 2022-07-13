using System.Xml.Serialization;

namespace Navisworks.Clash.Exporter.Setup.Manifest
{
    [XmlType(AnonymousType = true)]
    public partial class ComponentEntry
    {
        [XmlAttribute] public string AppName { get; set; }
        [XmlAttribute] public string AppType { get; set; }
        [XmlAttribute] public string ModuleName { get; set; }
    }
}