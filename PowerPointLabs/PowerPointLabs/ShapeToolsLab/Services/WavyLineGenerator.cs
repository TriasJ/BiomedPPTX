using System;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ShapeToolsLab.Services
{
    public static class WavyLineGenerator
    {
        public static PowerPoint.Shape Generate(
            PowerPoint.Slide slide,
            float startX,
            float startY,
            float width,
            float amplitude,
            float frequency,
            int points)
        {
            int segments = 256;
            if (points > 0 && points < segments)
            {
                segments = points;
            }

            float[,] polyPoints = new float[segments + 1, 2];
            float step = width / segments;

            for (int i = 0; i <= segments; i++)
            {
                polyPoints[i, 0] = startX + i * step;
                polyPoints[i, 1] = startY + amplitude * (float)Math.Sin(
                    frequency * i * Math.PI * 2.0 / segments);
            }

            PowerPoint.Shape shape = slide.Shapes.AddPolyline(polyPoints);
            shape.Line.ForeColor.RGB = 0x702000;
            shape.Line.Weight = 2;
            shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            return shape;
        }
    }
}
