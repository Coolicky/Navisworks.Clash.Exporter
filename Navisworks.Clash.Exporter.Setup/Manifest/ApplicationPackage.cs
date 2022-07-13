using System.Xml.Serialization;

namespace Navisworks.Clash.Exporter.Setup.Manifest
{
    [XmlType(AnonymousType = true)]
    [XmlRoot(Namespace = "", IsNullable = false)]
    public class ApplicationPackage
    {
        [XmlElement] public CompanyDetails CompanyDetails { get; set; }

        [XmlElement] public Components[] Components { get; set; }

        [XmlAttribute] public double SchemaVersion { get; set; }
        [XmlAttribute] public string AutodeskProduct { get; set; }
        [XmlAttribute] public string Name { get; set; }
        [XmlAttribute] public string Description { get; set; }
        [XmlAttribute] public string AppVersion { get; set; }
        [XmlAttribute] public string ProductCode { get; set; }
        [XmlAttribute] public string UpgradeCode { get; set; }
    }
}