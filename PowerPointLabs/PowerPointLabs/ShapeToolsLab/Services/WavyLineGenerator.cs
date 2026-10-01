using System;

using Microsoft.Office.Core;

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
            float step = width / points;
            PowerPoint.FreeformBuilder builder = slide.Shapes.BuildFreeform(
                MsoEditingType.msoEditingAuto, startX, startY);

            for (int i = 1; i <= points; i++)
            {
                float x = startX + i * step;
                float y = startY + amplitude * (float)Math.Sin(
                    frequency * i * Math.PI * 2.0 / points);
                builder.AddNodes(
                    MsoSegmentType.msoSegmentLine,
                    MsoEditingType.msoEditingAuto,
                    x, y);
            }

            PowerPoint.Shape shape = builder.ConvertToShape();
            shape.Line.ForeColor.RGB = 0x702000;
            shape.Line.Weight = 2;
            return shape;
        }
    }
}
