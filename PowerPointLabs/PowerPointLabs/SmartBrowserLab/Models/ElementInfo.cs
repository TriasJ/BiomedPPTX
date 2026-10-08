using System.Windows.Media;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class ElementInfo
    {
        public string Symbol { get; private set; }

        public double Radius { get; private set; }

        public Color Color { get; private set; }

        public ElementInfo(string symbol, double radius, Color color)
        {
            Symbol = symbol;
            Radius = radius;
            Color = color;
        }
    }
}
