using Autodesk.Navisworks.Api;

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
    }
}