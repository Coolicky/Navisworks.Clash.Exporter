using System.Xml.Serialization;

namespace Navisworks.Clash.Exporter.Setup.Manifest
{
    [XmlType(AnonymousType = true)]
    public class CompanyDetails
    {
        [XmlAttribute] public string Name { get; set; }
    }
}