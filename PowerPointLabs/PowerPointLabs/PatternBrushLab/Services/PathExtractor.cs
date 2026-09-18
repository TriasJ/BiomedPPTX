using System;
using System.Collections.Generic;
using System.Drawing;

using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PatternBrushLab.Services
{
    public class PathExtractor
    {
        private const int BezierSamples = 50;

        public List<PointF> ExtractPath(PowerPoint.Shape shape)
        {
            if (shape.Type == MsoShapeType.msoLine)
            {
                return ExtractLinePath(shape);
            }

            if (shape.Type == MsoShapeType.msoFreeform)
            {
                return ExtractFreeformPath(shape);
            }

            if (shape.Type == MsoShapeType.msoAutoShape)
            {
                return ExtractLinePath(shape);
            }

            return ExtractLinePath(shape);
        }

        private List<PointF> ExtractLinePath(PowerPoint.Shape shape)
        {
            float x1 = shape.Left;
            float y1 = shape.Top;
            float x2 = shape.Left + shape.Width;
            float y2 = shape.Top + shape.Height;

            if (shape.HorizontalFlip == MsoTriState.msoTrue)
            {
                float tmp = x1;
                x1 = x2;
                x2 = tmp;
            }
            if (shape.VerticalFlip == MsoTriState.msoTrue)
            {
                float tmp = y1;
                y1 = y2;
                y2 = tmp;
            }

            return new List<PointF> { new PointF(x1, y1), new PointF(x2, y2) };
        }

        private List<PointF> ExtractFreeformPath(PowerPoint.Shape shape)
        {
            var points = new List<PointF>();
            PowerPoint.ShapeNodes nodes = shape.Nodes;

            if (nodes.Count == 0)
            {
                return ExtractLinePath(shape);
            }

            float offsetX = shape.Left;
            float offsetY = shape.Top;

            for (int i = 1; i <= nodes.Count; i++)
            {
                PowerPoint.ShapeNode node = nodes[i];
                var pts = (object[,])node.Points;
                float nx = Convert.ToSingle(pts[1, 1]) + offsetX;
                float ny = Convert.ToSingle(pts[1, 2]) + offsetY;

                if (node.SegmentType == MsoSegmentType.msoSegmentCurve && i > 1)
                {
                    PointF startPt = points[points.Count - 1];
                    var controlPts = GetBezierControlPoints(nodes, i, offsetX, offsetY);
                    if (controlPts != null)
                    {
                        var curvePoints = SampleCubicBezier(
                            startPt, controlPts[0], controlPts[1], new PointF(nx, ny), BezierSamples);
                        for (int j = 1; j < curvePoints.Count; j++)
                        {
                            points.Add(curvePoints[j]);
                        }
                    }
                    else
                    {
                        points.Add(new PointF(nx, ny));
                    }
                }
                else
                {
                    if (points.Count == 0 || Distance(points[points.Count - 1], new PointF(nx, ny)) > 0.01f)
                    {
                        points.Add(new PointF(nx, ny));
                    }
                }
            }

            if (points.Count < 2)
            {
                return ExtractLinePath(shape);
            }

            return points;
        }

        private PointF[] GetBezierControlPoints(PowerPoint.ShapeNodes nodes, int nodeIndex,
            float offsetX, float offsetY)
        {
            try
            {
                PowerPoint.ShapeNode node = nodes[nodeIndex];
                var pts = (object[,])node.Points;

                int rows = pts.GetLength(0);
                if (rows >= 3)
                {
                    return new[]
                    {
                        new PointF(Convert.ToSingle(pts[1, 1]) + offsetX, Convert.ToSingle(pts[1, 2]) + offsetY),
                        new PointF(Convert.ToSingle(pts[2, 1]) + offsetX, Convert.ToSingle(pts[2, 2]) + offsetY)
                    };
                }
            }
            catch
            {
            }
            return null;
        }

        private List<PointF> SampleCubicBezier(PointF p0, PointF p1, PointF p2, PointF p3, int samples)
        {
            var result = new List<PointF>();
            for (int i = 0; i <= samples; i++)
            {
                float t = (float)i / samples;
                float u = 1f - t;
                float x = u * u * u * p0.X + 3 * u * u * t * p1.X + 3 * u * t * t * p2.X + t * t * t * p3.X;
                float y = u * u * u * p0.Y + 3 * u * u * t * p1.Y + 3 * u * t * t * p2.Y + t * t * t * p3.Y;
                result.Add(new PointF(x, y));
            }
            return result;
        }

        private float Distance(PointF a, PointF b)
        {
            float dx = a.X - b.X;
            float dy = a.Y - b.Y;
            return (float)Math.Sqrt(dx * dx + dy * dy);
        }
    }
}
