using System;

using Microsoft.Office.Core;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ShapeToolsLab.Services
{
    public static class GearStarGenerator
    {
        public static PowerPoint.Shape GenerateStar(
            PowerPoint.Slide slide,
            float cx,
            float cy,
            float outerRadius,
            float innerRadius,
            int pointCount)
        {
            int totalVertices = pointCount * 2;
            double startAngle = -Math.PI / 2.0;
            float firstX = cx + outerRadius * (float)Math.Cos(startAngle);
            float firstY = cy + outerRadius * (float)Math.Sin(startAngle);

            PowerPoint.FreeformBuilder builder = slide.Shapes.BuildFreeform(
                MsoEditingType.msoEditingAuto, firstX, firstY);

            for (int i = 1; i <= totalVertices; i++)
            {
                double angle = startAngle + Math.PI * i / pointCount;
                float r = (i % 2 == 0) ? outerRadius : innerRadius;
                float x = cx + r * (float)Math.Cos(angle);
                float y = cy + r * (float)Math.Sin(angle);
                builder.AddNodes(
                    MsoSegmentType.msoSegmentLine,
                    MsoEditingType.msoEditingAuto,
                    x, y);
            }

            PowerPoint.Shape shape = builder.ConvertToShape();
            shape.Fill.ForeColor.RGB = 0x40A0D0;
            shape.Line.ForeColor.RGB = 0x206080;
            shape.Line.Weight = 1.5f;
            return shape;
        }

        public static PowerPoint.Shape GenerateGear(
            PowerPoint.Slide slide,
            float cx,
            float cy,
            float outerRadius,
            float innerRadius,
            int toothCount)
        {
            float toothAngle = (float)(2.0 * Math.PI / toothCount);
            float toothWidthFraction = 0.35f;

            float firstAngle = 0f;
            float firstX = cx + innerRadius * (float)Math.Cos(firstAngle);
            float firstY = cy + innerRadius * (float)Math.Sin(firstAngle);

            PowerPoint.FreeformBuilder builder = slide.Shapes.BuildFreeform(
                MsoEditingType.msoEditingAuto, firstX, firstY);

            for (int i = 0; i < toothCount; i++)
            {
                float baseAngle = i * toothAngle;
                float a1 = baseAngle + toothAngle * (0.5f - toothWidthFraction / 2f);
                float a2 = baseAngle + toothAngle * (0.5f + toothWidthFraction / 2f);
                float a3 = baseAngle + toothAngle;

                AddPoint(builder, cx, cy, outerRadius, a1);
                AddPoint(builder, cx, cy, outerRadius, a2);
                AddPoint(builder, cx, cy, innerRadius, a3);
            }

            PowerPoint.Shape shape = builder.ConvertToShape();
            shape.Fill.ForeColor.RGB = 0x808080;
            shape.Line.ForeColor.RGB = 0x404040;
            shape.Line.Weight = 1.5f;
            return shape;
        }

        private static void AddPoint(
            PowerPoint.FreeformBuilder builder,
            float cx,
            float cy,
            float radius,
            float angle)
        {
            float x = cx + radius * (float)Math.Cos(angle);
            float y = cy + radius * (float)Math.Sin(angle);
            builder.AddNodes(
                MsoSegmentType.msoSegmentLine,
                MsoEditingType.msoEditingAuto,
                x, y);
        }
    }
}
