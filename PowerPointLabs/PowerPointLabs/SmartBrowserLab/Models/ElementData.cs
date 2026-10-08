using System.Collections.Generic;
using System.Windows.Media;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public static class ElementData
    {
        private static readonly Dictionary<int, ElementInfo> Elements = new Dictionary<int, ElementInfo>
        {
            { 1,  new ElementInfo("H",  0.25, Color.FromRgb(224, 224, 224)) },
            { 6,  new ElementInfo("C",  0.40, Color.FromRgb(64, 64, 64)) },
            { 7,  new ElementInfo("N",  0.38, Color.FromRgb(32, 32, 255)) },
            { 8,  new ElementInfo("O",  0.36, Color.FromRgb(255, 32, 32)) },
            { 9,  new ElementInfo("F",  0.35, Color.FromRgb(144, 224, 80)) },
            { 15, new ElementInfo("P",  0.45, Color.FromRgb(255, 128, 0)) },
            { 16, new ElementInfo("S",  0.45, Color.FromRgb(255, 255, 48)) },
            { 17, new ElementInfo("Cl", 0.43, Color.FromRgb(32, 255, 32)) },
            { 35, new ElementInfo("Br", 0.45, Color.FromRgb(166, 41, 41)) },
            { 53, new ElementInfo("I",  0.50, Color.FromRgb(148, 0, 148)) },
            { 26, new ElementInfo("Fe", 0.40, Color.FromRgb(224, 102, 51)) },
            { 30, new ElementInfo("Zn", 0.39, Color.FromRgb(125, 128, 176)) },
            { 20, new ElementInfo("Ca", 0.48, Color.FromRgb(61, 255, 0)) },
            { 11, new ElementInfo("Na", 0.50, Color.FromRgb(171, 92, 242)) },
            { 19, new ElementInfo("K",  0.55, Color.FromRgb(143, 64, 212)) },
            { 12, new ElementInfo("Mg", 0.45, Color.FromRgb(138, 255, 0)) },
        };

        private static readonly ElementInfo DefaultElement =
            new ElementInfo("?", 0.40, Color.FromRgb(255, 105, 180));

        public static ElementInfo GetElement(int atomicNumber)
        {
            ElementInfo info;
            if (Elements.TryGetValue(atomicNumber, out info))
            {
                return info;
            }

            return DefaultElement;
        }

        public static string GetSymbol(int atomicNumber)
        {
            return GetElement(atomicNumber).Symbol;
        }

        public static double GetRadius(int atomicNumber)
        {
            return GetElement(atomicNumber).Radius;
        }

        public static Color GetColor(int atomicNumber)
        {
            return GetElement(atomicNumber).Color;
        }

        public static string GetHexColor(int atomicNumber)
        {
            Color c = GetColor(atomicNumber);
            return string.Format("#{0:X2}{1:X2}{2:X2}", c.R, c.G, c.B);
        }
    }
}
