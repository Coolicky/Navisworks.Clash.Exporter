using System.Xml.Serialization;

namespace Navisworks.Clash.Exporter.Setup.Manifest
{
    [XmlType(AnonymousType = true)]
    public partial class RuntimeRequirements
    {
        [XmlAttribute] public string OS { get; set; }
        [XmlAttribute] public string Platform { get; set; }
        [XmlAttribute] public string SeriesMin { get; set; }
        [XmlAttribute] public string SeriesMax { get; set; }
    }
}