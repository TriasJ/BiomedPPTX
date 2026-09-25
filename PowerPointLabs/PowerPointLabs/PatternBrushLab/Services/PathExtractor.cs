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
                return ExtractAutoShapePath(shape);
            }

            return ExtractLinePath(shape);
        }

        private List<PointF> ExtractAutoShapePath(PowerPoint.Shape shape)
        {
            PowerPoint.Slide slide = null;
            try
            {
                slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide as PowerPoint.Slide;
            }
            catch (Exception)
            {
            }

            if (slide == null)
            {
                return GenerateOutlinePath(shape);
            }

            PowerPoint.Shape duplicate = null;
            try
            {
                shape.Copy();
                System.Threading.Thread.Sleep(150);
                slide.Shapes.Paste();
                System.Threading.Thread.Sleep(100);
                duplicate = slide.Shapes[slide.Shapes.Count];

                dynamic dynDup = duplicate;
                PowerPoint.Shape freeform = null;
                try
                {
                    freeform = dynDup.ConvertToFreeform();
                    System.Threading.Thread.Sleep(100);
                }
                catch (Exception)
                {
                }

                PowerPoint.Shape shapeToExtract = freeform != null ? freeform : duplicate;

                List<PointF> path;
                if (shapeToExtract.Type == MsoShapeType.msoFreeform)
                {
                    path = ExtractFreeformPath(shapeToExtract);
                }
                else
                {
                    path = GenerateOutlinePath(shapeToExtract);
                }

                if (path.Count > 2)
                {
                    float dist = Distance(path[0], path[path.Count - 1]);
                    if (dist > 1.0f)
                    {
                        path.Add(path[0]);
                    }
                }

                try
                {
                    shapeToExtract.Delete();
                }
                catch (Exception)
                {
                    try
                    {
                        if (duplicate != null)
                        {
                            duplicate.Delete();
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                return path;
            }
            catch (Exception)
            {
                if (duplicate != null)
                {
                    try
                    {
                        duplicate.Delete();
                    }
                    catch (Exception)
                    {
                    }
                }

                return GenerateOutlinePath(shape);
            }
        }

        private List<PointF> GenerateOutlinePath(PowerPoint.Shape shape)
        {
            float cx = shape.Left + shape.Width / 2;
            float cy = shape.Top + shape.Height / 2;
            float rx = shape.Width / 2;
            float ry = shape.Height / 2;

            var points = new List<PointF>();
            int segments = 72;

            for (int i = 0; i <= segments; i++)
            {
                double angle = 2.0 * Math.PI * i / segments;
                float x = cx + rx * (float)Math.Cos(angle);
                float y = cy + ry * (float)Math.Sin(angle);
                points.Add(new PointF(x, y));
            }

            return points;
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
                PointF nodePoint = GetNodePoint(node, offsetX, offsetY);
                float nx = nodePoint.X;
                float ny = nodePoint.Y;

                if (node.SegmentType == MsoSegmentType.msoSegmentCurve && i > 1)
                {
                    PointF startPt = points[points.Count - 1];
                    PointF[] controlPts = GetBezierControlPoints(nodes, i, offsetX, offsetY);
                    if (controlPts != null)
                    {
                        List<PointF> curvePoints = SampleCubicBezier(
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

        private PointF GetNodePoint(PowerPoint.ShapeNode node, float offsetX, float offsetY)
        {
            try
            {
                dynamic pts = node.Points;
                float x = Convert.ToSingle(pts[1, 1]) + offsetX;
                float y = Convert.ToSingle(pts[1, 2]) + offsetY;
                return new PointF(x, y);
            }
            catch (Exception)
            {
                return new PointF(offsetX, offsetY);
            }
        }

        private PointF[] GetBezierControlPoints(PowerPoint.ShapeNodes nodes, int nodeIndex,
            float offsetX, float offsetY)
        {
            try
            {
                PowerPoint.ShapeNode node = nodes[nodeIndex];
                dynamic pts = node.Points;

                Array arr = pts as Array;
                if (arr != null && arr.GetLength(0) >= 3)
                {
                    return new[]
                    {
                        new PointF(Convert.ToSingle(pts[1, 1]) + offsetX, Convert.ToSingle(pts[1, 2]) + offsetY),
                        new PointF(Convert.ToSingle(pts[2, 1]) + offsetX, Convert.ToSingle(pts[2, 2]) + offsetY)
                    };
                }
            }
            catch (Exception)
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
