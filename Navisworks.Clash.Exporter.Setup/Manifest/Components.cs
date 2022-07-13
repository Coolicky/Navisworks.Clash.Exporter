using System.Xml.Serialization;

namespace Navisworks.Clash.Exporter.Setup.Manifest
{
    [XmlType(AnonymousType = true)]
    public partial class Components
    {
        [XmlElement] public RuntimeRequirements RuntimeRequirements { get; set; }
        [XmlElement] public ComponentEntry ComponentEntry { get; set; }
        [XmlAttribute] public string Description { get; set; }
    }
}