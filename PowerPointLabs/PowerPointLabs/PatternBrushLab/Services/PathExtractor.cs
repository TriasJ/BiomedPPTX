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
            float left = shape.Left;
            float top = shape.Top;
            float w = shape.Width;
            float h = shape.Height;
            float cx = left + w / 2;
            float cy = top + h / 2;

            try
            {
                MsoAutoShapeType shapeType = shape.AutoShapeType;

                if (shapeType == MsoAutoShapeType.msoShapeIsoscelesTriangle)
                {
                    return InterpolatePolygon(new List<PointF>
                    {
                        new PointF(cx, top),
                        new PointF(left + w, top + h),
                        new PointF(left, top + h),
                        new PointF(cx, top)
                    }, 10);
                }

                if (shapeType == MsoAutoShapeType.msoShapeRightTriangle)
                {
                    return InterpolatePolygon(new List<PointF>
                    {
                        new PointF(left, top),
                        new PointF(left, top + h),
                        new PointF(left + w, top + h),
                        new PointF(left, top)
                    }, 10);
                }

                if (shapeType == MsoAutoShapeType.msoShapeRectangle ||
                    shapeType == MsoAutoShapeType.msoShapeRoundedRectangle)
                {
                    return InterpolatePolygon(new List<PointF>
                    {
                        new PointF(left, top),
                        new PointF(left + w, top),
                        new PointF(left + w, top + h),
                        new PointF(left, top + h),
                        new PointF(left, top)
                    }, 10);
                }

                if (shapeType == MsoAutoShapeType.msoShapeDiamond)
                {
                    return InterpolatePolygon(new List<PointF>
                    {
                        new PointF(cx, top),
                        new PointF(left + w, cy),
                        new PointF(cx, top + h),
                        new PointF(left, cy),
                        new PointF(cx, top)
                    }, 10);
                }

                if (shapeType == MsoAutoShapeType.msoShapePentagon ||
                    shapeType == MsoAutoShapeType.msoShapeRegularPentagon)
                {
                    return GenerateRegularPolygon(cx, cy, w / 2, h / 2, 5, -Math.PI / 2);
                }

                if (shapeType == MsoAutoShapeType.msoShapeHexagon)
                {
                    return GenerateRegularPolygon(cx, cy, w / 2, h / 2, 6, 0);
                }

                if (shapeType == MsoAutoShapeType.msoShapeOctagon)
                {
                    return GenerateRegularPolygon(cx, cy, w / 2, h / 2, 8, Math.PI / 8);
                }

                if (shapeType == MsoAutoShapeType.msoShape4pointStar)
                {
                    return GenerateStarPath(cx, cy, w / 2, h / 2, 4);
                }

                if (shapeType == MsoAutoShapeType.msoShape5pointStar)
                {
                    return GenerateStarPath(cx, cy, w / 2, h / 2, 5);
                }

                if (shapeType == MsoAutoShapeType.msoShape6pointStar)
                {
                    return GenerateStarPath(cx, cy, w / 2, h / 2, 6);
                }

                if (shapeType == MsoAutoShapeType.msoShape8pointStar)
                {
                    return GenerateStarPath(cx, cy, w / 2, h / 2, 8);
                }

                if (shapeType == MsoAutoShapeType.msoShape10pointStar)
                {
                    return GenerateStarPath(cx, cy, w / 2, h / 2, 10);
                }

                if (shapeType == MsoAutoShapeType.msoShape12pointStar)
                {
                    return GenerateStarPath(cx, cy, w / 2, h / 2, 12);
                }

                if (shapeType == MsoAutoShapeType.msoShape16pointStar)
                {
                    return GenerateStarPath(cx, cy, w / 2, h / 2, 16);
                }

                if (shapeType == MsoAutoShapeType.msoShape24pointStar)
                {
                    return GenerateStarPath(cx, cy, w / 2, h / 2, 24);
                }

                if (shapeType == MsoAutoShapeType.msoShape32pointStar)
                {
                    return GenerateStarPath(cx, cy, w / 2, h / 2, 32);
                }

                if (shapeType == MsoAutoShapeType.msoShapeCross)
                {
                    float arm = w / 3;
                    return InterpolatePolygon(new List<PointF>
                    {
                        new PointF(left + arm, top),
                        new PointF(left + 2 * arm, top),
                        new PointF(left + 2 * arm, top + arm),
                        new PointF(left + w, top + arm),
                        new PointF(left + w, top + 2 * arm),
                        new PointF(left + 2 * arm, top + 2 * arm),
                        new PointF(left + 2 * arm, top + h),
                        new PointF(left + arm, top + h),
                        new PointF(left + arm, top + 2 * arm),
                        new PointF(left, top + 2 * arm),
                        new PointF(left, top + arm),
                        new PointF(left + arm, top + arm),
                        new PointF(left + arm, top)
                    }, 5);
                }
            }
            catch (Exception)
            {
            }

            return GenerateEllipsePath(cx, cy, w / 2, h / 2);
        }

        private List<PointF> GenerateEllipsePath(float cx, float cy, float rx, float ry)
        {
            var points = new List<PointF>();
            int segments = 72;
            for (int i = 0; i <= segments; i++)
            {
                double angle = 2.0 * Math.PI * i / segments;
                points.Add(new PointF(cx + rx * (float)Math.Cos(angle), cy + ry * (float)Math.Sin(angle)));
            }

            return points;
        }

        private List<PointF> GenerateRegularPolygon(float cx, float cy, float rx, float ry, int sides, double startAngle)
        {
            var vertices = new List<PointF>();
            for (int i = 0; i <= sides; i++)
            {
                double angle = startAngle + 2.0 * Math.PI * i / sides;
                vertices.Add(new PointF(cx + rx * (float)Math.Cos(angle), cy + ry * (float)Math.Sin(angle)));
            }

            return InterpolatePolygon(vertices, 10);
        }

        private List<PointF> GenerateStarPath(float cx, float cy, float rx, float ry, int numPoints)
        {
            var vertices = new List<PointF>();
            float innerRx = rx * 0.4f;
            float innerRy = ry * 0.4f;
            double startAngle = -Math.PI / 2;

            for (int i = 0; i <= numPoints * 2; i++)
            {
                double angle = startAngle + Math.PI * i / numPoints;
                bool isOuter = (i % 2 == 0);
                float rX = isOuter ? rx : innerRx;
                float rY = isOuter ? ry : innerRy;
                vertices.Add(new PointF(cx + rX * (float)Math.Cos(angle), cy + rY * (float)Math.Sin(angle)));
            }

            return InterpolatePolygon(vertices, 5);
        }

        private List<PointF> InterpolatePolygon(List<PointF> vertices, int pointsPerEdge)
        {
            var result = new List<PointF>();
            for (int i = 0; i < vertices.Count - 1; i++)
            {
                for (int j = 0; j < pointsPerEdge; j++)
                {
                    float t = (float)j / pointsPerEdge;
                    float x = vertices[i].X + t * (vertices[i + 1].X - vertices[i].X);
                    float y = vertices[i].Y + t * (vertices[i + 1].Y - vertices[i].Y);
                    result.Add(new PointF(x, y));
                }
            }

            result.Add(vertices[vertices.Count - 1]);
            return result;
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
