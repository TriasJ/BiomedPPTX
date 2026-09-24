using System;
using System.Collections.Generic;
using System.Drawing;

namespace PowerPointLabs.PatternBrushLab.Services
{
    public struct TilePlacement
    {
        public PointF Position;
        public float RotationDegrees;
        public int Index;
    }

    public class TilingEngine
    {
        private Random _rng;

        public List<TilePlacement> ComputePlacements(
            List<PointF> pathPoints,
            float tileWidth,
            float overlapPt,
            float scatterPt = 0,
            float rotationJitterDeg = 0,
            float rotationOffset = 0,
            int seed = 42)
        {
            if (pathPoints == null || pathPoints.Count < 2 || tileWidth <= 0)
            {
                return new List<TilePlacement>();
            }

            _rng = new Random(seed);

            float step = tileWidth - overlapPt;
            if (step <= 0)
            {
                step = tileWidth * 0.5f;
            }

            var arcLengths = ComputeArcLengths(pathPoints);
            float totalLength = arcLengths[arcLengths.Count - 1];

            int tileCount = Math.Max(1, (int)Math.Ceiling(totalLength / step) + 1);

            var placements = new List<TilePlacement>();
            for (int i = 0; i < tileCount; i++)
            {
                float distance = i * step;
                if (distance > totalLength)
                {
                    break;
                }

                var result = SampleAtDistance(pathPoints, arcLengths, distance);
                PointF pos = result.Item1;
                float tangentAngle = result.Item2;

                if (scatterPt > 0)
                {
                    float perpRad = (float)((tangentAngle + 90) * Math.PI / 180.0);
                    float offset = (float)(_rng.NextDouble() * 2 - 1) * scatterPt;
                    pos = new PointF(
                        pos.X + offset * (float)Math.Cos(perpRad),
                        pos.Y + offset * (float)Math.Sin(perpRad));
                }

                float finalRotation = tangentAngle + rotationOffset;
                if (rotationJitterDeg > 0)
                {
                    finalRotation += (float)(_rng.NextDouble() * 2 - 1) * rotationJitterDeg;
                }

                placements.Add(new TilePlacement
                {
                    Position = pos,
                    RotationDegrees = finalRotation,
                    Index = i
                });
            }

            return placements;
        }

        public List<TilePlacement> ComputeGridPlacements(
            RectangleF area,
            float tileWidth,
            float tileHeight,
            float overlapX,
            float overlapY,
            bool offsetRows,
            float offsetAmount)
        {
            float stepX = tileWidth - overlapX;
            float stepY = tileHeight - overlapY;
            if (stepX <= 0)
            {
                stepX = tileWidth * 0.5f;
            }

            if (stepY <= 0)
            {
                stepY = tileHeight * 0.5f;
            }

            int cols = (int)Math.Ceiling(area.Width / stepX) + 1;
            int rows = (int)Math.Ceiling(area.Height / stepY) + 1;

            var placements = new List<TilePlacement>();
            for (int r = 0; r < rows; r++)
            {
                float rowOffset = (offsetRows && r % 2 == 1) ? offsetAmount : 0;

                for (int c = 0; c < cols; c++)
                {
                    placements.Add(new TilePlacement
                    {
                        Position = new PointF(
                            area.X + c * stepX + rowOffset,
                            area.Y + r * stepY),
                        RotationDegrees = 0,
                        Index = placements.Count
                    });
                }
            }
            return placements;
        }

        private List<float> ComputeArcLengths(List<PointF> points)
        {
            var lengths = new List<float> { 0f };
            for (int i = 1; i < points.Count; i++)
            {
                float dx = points[i].X - points[i - 1].X;
                float dy = points[i].Y - points[i - 1].Y;
                float segLen = (float)Math.Sqrt(dx * dx + dy * dy);
                lengths.Add(lengths[i - 1] + segLen);
            }
            return lengths;
        }

        private Tuple<PointF, float> SampleAtDistance(
            List<PointF> points,
            List<float> arcLengths,
            float targetDist)
        {
            int lo = 0, hi = arcLengths.Count - 1;
            while (lo < hi - 1)
            {
                int mid = (lo + hi) / 2;
                if (arcLengths[mid] <= targetDist)
                {
                    lo = mid;
                }
                else
                {
                    hi = mid;
                }
            }

            int j = lo;
            float segStart = arcLengths[j];
            float segEnd = arcLengths[j + 1];
            float segLen = segEnd - segStart;

            float t = segLen > 0 ? (targetDist - segStart) / segLen : 0;
            t = Math.Max(0, Math.Min(1, t));

            float px = points[j].X + t * (points[j + 1].X - points[j].X);
            float py = points[j].Y + t * (points[j + 1].Y - points[j].Y);

            float dx = points[j + 1].X - points[j].X;
            float dy = points[j + 1].Y - points[j].Y;
            float angleDeg = (float)(Math.Atan2(dy, dx) * 180.0 / Math.PI);

            return Tuple.Create(new PointF(px, py), angleDeg);
        }
    }
}
