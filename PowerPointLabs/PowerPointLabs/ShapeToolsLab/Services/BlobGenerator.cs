using System;
using System.Collections.Generic;
using System.Drawing;

using Microsoft.Office.Core;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ShapeToolsLab.Services
{
    public static class BlobGenerator
    {
        public static PowerPoint.Shape GenerateBadge(
            PowerPoint.Slide slide,
            float cx,
            float cy,
            float baseRadius,
            float bumpHeight,
            int bumpCount,
            int segments)
        {
            if (segments < 64)
            {
                segments = 256;
            }

            float[,] points = new float[segments + 1, 2];

            for (int i = 0; i <= segments; i++)
            {
                double angle = 2.0 * Math.PI * i / segments;
                float r = baseRadius + (float)Math.Cos(angle * bumpCount) * bumpHeight;
                points[i, 0] = cx + r * (float)Math.Sin(angle);
                points[i, 1] = cy + r * -(float)Math.Cos(angle);
            }

            PowerPoint.Shape shape = slide.Shapes.AddPolyline((object)points);
            shape.Fill.ForeColor.RGB = 0x4090D0;
            shape.Line.Visible = MsoTriState.msoFalse;
            return shape;
        }

        public static PowerPoint.Shape Generate(
            PowerPoint.Slide slide,
            float cx,
            float cy,
            float baseRadius,
            float jitter,
            int pointCount,
            int seed)
        {
            Random rng = new Random(seed);
            int densityMultiplier = 6;
            int totalPoints = pointCount * densityMultiplier;
            List<PointF> rawPoints = new List<PointF>();

            for (int i = 0; i < pointCount; i++)
            {
                double angle = 2.0 * Math.PI * i / pointCount;
                float r = baseRadius + (float)(rng.NextDouble() * 2.0 - 1.0) * jitter;
                rawPoints.Add(new PointF(
                    cx + r * (float)Math.Cos(angle),
                    cy + r * (float)Math.Sin(angle)));
            }

            List<PointF> smooth = Interpolate(rawPoints, totalPoints);

            PowerPoint.FreeformBuilder builder = slide.Shapes.BuildFreeform(
                MsoEditingType.msoEditingAuto, smooth[0].X, smooth[0].Y);

            for (int i = 1; i < smooth.Count; i++)
            {
                builder.AddNodes(
                    MsoSegmentType.msoSegmentLine,
                    MsoEditingType.msoEditingAuto,
                    smooth[i].X, smooth[i].Y);
            }

            builder.AddNodes(
                MsoSegmentType.msoSegmentLine,
                MsoEditingType.msoEditingAuto,
                smooth[0].X, smooth[0].Y);

            PowerPoint.Shape shape = builder.ConvertToShape();
            shape.Fill.ForeColor.RGB = 0xD09040;
            shape.Line.ForeColor.RGB = 0x804020;
            shape.Line.Weight = 1.5f;
            return shape;
        }

        private static List<PointF> Interpolate(List<PointF> control, int outputCount)
        {
            List<PointF> result = new List<PointF>();
            int n = control.Count;

            for (int i = 0; i < outputCount; i++)
            {
                float t = (float)i / outputCount * n;
                int idx = (int)t;
                float frac = t - idx;

                PointF p0 = control[((idx - 1) % n + n) % n];
                PointF p1 = control[idx % n];
                PointF p2 = control[(idx + 1) % n];
                PointF p3 = control[(idx + 2) % n];

                float x = CatmullRom(p0.X, p1.X, p2.X, p3.X, frac);
                float y = CatmullRom(p0.Y, p1.Y, p2.Y, p3.Y, frac);
                result.Add(new PointF(x, y));
            }

            return result;
        }

        private static float CatmullRom(float p0, float p1, float p2, float p3, float t)
        {
            float t2 = t * t;
            float t3 = t2 * t;
            return 0.5f * (
                (2f * p1) +
                (-p0 + p2) * t +
                (2f * p0 - 5f * p1 + 4f * p2 - p3) * t2 +
                (-p0 + 3f * p1 - 3f * p2 + p3) * t3);
        }
    }
}
