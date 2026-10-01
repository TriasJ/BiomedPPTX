using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text;

using Microsoft.Office.Core;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ShapeToolsLab.Services
{
    public static class MotionPathConverter
    {
        public static void Convert(
            PowerPoint.Shape targetShape,
            PowerPoint.Shape pathShape,
            PowerPoint.Slide slide)
        {
            List<PointF> points = ExtractPoints(pathShape);
            if (points.Count < 2)
            {
                throw new InvalidOperationException("Path shape must have at least 2 points.");
            }

            float slideW = slide.CustomLayout.Width;
            float slideH = slide.CustomLayout.Height;

            float startX = points[0].X;
            float startY = points[0].Y;

            StringBuilder pathStr = new StringBuilder();
            pathStr.Append("M 0.0000 0.0000 ");

            for (int i = 1; i < points.Count; i++)
            {
                float dx = (points[i].X - startX) / slideW;
                float dy = (points[i].Y - startY) / slideH;
                pathStr.AppendFormat(CultureInfo.InvariantCulture, "L {0:F4} {1:F4} ", dx, dy);
            }

            pathStr.Append("E");

            PowerPoint.Effect effect = slide.TimeLine.MainSequence.AddEffect(
                targetShape,
                PowerPoint.MsoAnimEffect.msoAnimEffectPathDown,
                PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);

            effect.Behaviors[1].MotionEffect.Path = pathStr.ToString();

            float targetCX = targetShape.Left + targetShape.Width / 2f;
            float targetCY = targetShape.Top + targetShape.Height / 2f;
            targetShape.Left = targetShape.Left + (startX - targetCX);
            targetShape.Top = targetShape.Top + (startY - targetCY);
        }

        private static List<PointF> ExtractPoints(PowerPoint.Shape shape)
        {
            List<PointF> points = new List<PointF>();

            if (shape.Type == MsoShapeType.msoLine)
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

                points.Add(new PointF(x1, y1));
                points.Add(new PointF(x2, y2));
            }
            else if (shape.Type == MsoShapeType.msoFreeform)
            {
                float offsetX = shape.Left;
                float offsetY = shape.Top;

                for (int i = 1; i <= shape.Nodes.Count; i++)
                {
                    try
                    {
                        dynamic pts = shape.Nodes[i].Points;
                        float x = System.Convert.ToSingle(pts[1, 1]) + offsetX;
                        float y = System.Convert.ToSingle(pts[1, 2]) + offsetY;
                        points.Add(new PointF(x, y));
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            else
            {
                points.Add(new PointF(shape.Left, shape.Top));
                points.Add(new PointF(shape.Left + shape.Width, shape.Top + shape.Height));
            }

            return points;
        }
    }
}
