using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public static class MoleculeSvgRenderer
    {
        private const double DefaultPadding = 40.0;
        private const string BackgroundColor = "#1a1a2e";
        private const string BondColor = "#c8c8c8";
        private const double BondStrokeWidth = 1.8;
        private const double DoubleOffset = 2.5;
        private const double TripleOffset = 3.5;
        private const double HeteroBackgroundRadius = 10.0;
        private const double FontSize = 13.0;

        public static string RenderSvg(PubChemCompoundData compound, int width, int height)
        {
            if (width <= 0)
            {
                width = 400;
            }

            if (height <= 0)
            {
                height = 400;
            }

            if (!HasValidStructureData(compound))
            {
                return BuildNoStructureSvg(width, height);
            }

            PubChemCoords coords = compound.Coords[0];
            PubChemConformer conformer = coords.Conformers[0];

            Dictionary<int, double> xMap = new Dictionary<int, double>();
            Dictionary<int, double> yMap = new Dictionary<int, double>();
            BuildCoordinateMaps(coords, conformer, xMap, yMap);

            if (xMap.Count == 0)
            {
                return BuildNoStructureSvg(width, height);
            }

            Dictionary<int, int> atomElementMap = BuildAtomElementMap(compound.Atoms);

            double minX;
            double maxX;
            double minY;
            double maxY;
            ComputeBounds(xMap, yMap, out minX, out maxX, out minY, out maxY);

            double rangeX = maxX - minX;
            double rangeY = maxY - minY;
            if (rangeX < 0.001)
            {
                rangeX = 1.0;
            }

            if (rangeY < 0.001)
            {
                rangeY = 1.0;
            }

            double drawWidth = width - (2.0 * DefaultPadding);
            double drawHeight = height - (2.0 * DefaultPadding);
            double scaleX = drawWidth / rangeX;
            double scaleY = drawHeight / rangeY;
            double scale = Math.Min(scaleX, scaleY);

            double offsetX = DefaultPadding + ((drawWidth - (rangeX * scale)) / 2.0);
            double offsetY = DefaultPadding + ((drawHeight - (rangeY * scale)) / 2.0);

            HashSet<int> hiddenHydrogens = FindHiddenHydrogens(compound, atomElementMap);

            StringBuilder sb = new StringBuilder();
            AppendSvgHeader(sb, width, height);
            AppendBonds(sb, compound.Bonds, xMap, yMap, minX, minY, maxY, scale, offsetX, offsetY, hiddenHydrogens);
            AppendAtoms(sb, compound.Atoms, xMap, yMap, atomElementMap, minX, minY, maxY, scale, offsetX, offsetY, hiddenHydrogens);
            AppendSvgFooter(sb);

            return sb.ToString();
        }

        public static string RenderSvg(PubChemCompoundData compound)
        {
            return RenderSvg(compound, 400, 400);
        }

        private static bool HasValidStructureData(PubChemCompoundData compound)
        {
            if (compound == null)
            {
                return false;
            }

            if (compound.Atoms == null || compound.Atoms.AtomIds == null || compound.Atoms.Elements == null)
            {
                return false;
            }

            if (compound.Coords == null || compound.Coords.Length == 0)
            {
                return false;
            }

            PubChemCoords coords = compound.Coords[0];
            if (coords == null || coords.Conformers == null || coords.Conformers.Length == 0)
            {
                return false;
            }

            PubChemConformer conformer = coords.Conformers[0];
            if (conformer == null || conformer.X == null || conformer.Y == null)
            {
                return false;
            }

            if (coords.AtomIds == null || coords.AtomIds.Length == 0)
            {
                return false;
            }

            return true;
        }

        private static string BuildNoStructureSvg(int width, int height)
        {
            StringBuilder sb = new StringBuilder();
            AppendSvgHeader(sb, width, height);
            sb.Append(string.Format(
                CultureInfo.InvariantCulture,
                "  <text x=\"{0}\" y=\"{1}\" text-anchor=\"middle\" font-family=\"Arial, sans-serif\" font-size=\"16\" fill=\"#999999\">No structure available</text>\n",
                width / 2.0,
                height / 2.0));
            AppendSvgFooter(sb);
            return sb.ToString();
        }

        private static void BuildCoordinateMaps(
            PubChemCoords coords,
            PubChemConformer conformer,
            Dictionary<int, double> xMap,
            Dictionary<int, double> yMap)
        {
            int count = Math.Min(coords.AtomIds.Length, Math.Min(conformer.X.Length, conformer.Y.Length));
            for (int i = 0; i < count; i++)
            {
                int atomId = coords.AtomIds[i];
                xMap[atomId] = conformer.X[i];
                yMap[atomId] = conformer.Y[i];
            }
        }

        private static Dictionary<int, int> BuildAtomElementMap(PubChemAtoms atoms)
        {
            Dictionary<int, int> map = new Dictionary<int, int>();
            if (atoms == null || atoms.AtomIds == null || atoms.Elements == null)
            {
                return map;
            }

            int count = Math.Min(atoms.AtomIds.Length, atoms.Elements.Length);
            for (int i = 0; i < count; i++)
            {
                map[atoms.AtomIds[i]] = atoms.Elements[i];
            }

            return map;
        }

        private static void ComputeBounds(
            Dictionary<int, double> xMap,
            Dictionary<int, double> yMap,
            out double minX,
            out double maxX,
            out double minY,
            out double maxY)
        {
            minX = double.MaxValue;
            maxX = double.MinValue;
            minY = double.MaxValue;
            maxY = double.MinValue;

            foreach (KeyValuePair<int, double> entry in xMap)
            {
                double x = entry.Value;
                if (x < minX)
                {
                    minX = x;
                }

                if (x > maxX)
                {
                    maxX = x;
                }
            }

            foreach (KeyValuePair<int, double> entry in yMap)
            {
                double y = entry.Value;
                if (y < minY)
                {
                    minY = y;
                }

                if (y > maxY)
                {
                    maxY = y;
                }
            }
        }

        private static HashSet<int> FindHiddenHydrogens(
            PubChemCompoundData compound,
            Dictionary<int, int> atomElementMap)
        {
            HashSet<int> hidden = new HashSet<int>();
            PubChemBonds bonds = compound.Bonds;
            if (bonds == null || bonds.AtomId1 == null || bonds.AtomId2 == null)
            {
                return hidden;
            }

            int bondCount = Math.Min(bonds.AtomId1.Length, bonds.AtomId2.Length);
            for (int i = 0; i < bondCount; i++)
            {
                int a1 = bonds.AtomId1[i];
                int a2 = bonds.AtomId2[i];

                int elem1;
                int elem2;
                bool hasElem1 = atomElementMap.TryGetValue(a1, out elem1);
                bool hasElem2 = atomElementMap.TryGetValue(a2, out elem2);

                if (hasElem1 && hasElem2)
                {
                    if (elem1 == 1 && elem2 != 1)
                    {
                        hidden.Add(a1);
                    }
                    else if (elem2 == 1 && elem1 != 1)
                    {
                        hidden.Add(a2);
                    }
                }
            }

            return hidden;
        }

        private static void AppendSvgHeader(StringBuilder sb, int width, int height)
        {
            sb.Append(string.Format(
                CultureInfo.InvariantCulture,
                "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{0}\" height=\"{1}\" viewBox=\"0 0 {0} {1}\">\n",
                width,
                height));
            sb.Append(string.Format(
                CultureInfo.InvariantCulture,
                "  <rect width=\"100%\" height=\"100%\" fill=\"{0}\"/>\n",
                BackgroundColor));
        }

        private static void AppendSvgFooter(StringBuilder sb)
        {
            sb.Append("</svg>");
        }

        private static void AppendBonds(
            StringBuilder sb,
            PubChemBonds bonds,
            Dictionary<int, double> xMap,
            Dictionary<int, double> yMap,
            double minX,
            double minY,
            double maxY,
            double scale,
            double offsetX,
            double offsetY,
            HashSet<int> hiddenHydrogens)
        {
            if (bonds == null || bonds.AtomId1 == null || bonds.AtomId2 == null)
            {
                return;
            }

            int bondCount = Math.Min(bonds.AtomId1.Length, bonds.AtomId2.Length);
            int[] orderArray = bonds.Order;

            for (int i = 0; i < bondCount; i++)
            {
                int a1 = bonds.AtomId1[i];
                int a2 = bonds.AtomId2[i];

                if (hiddenHydrogens.Contains(a1) || hiddenHydrogens.Contains(a2))
                {
                    continue;
                }

                double rawX1 = 0;
                double rawY1 = 0;
                double rawX2 = 0;
                double rawY2 = 0;
                bool hasA1 = xMap.TryGetValue(a1, out rawX1) && yMap.TryGetValue(a1, out rawY1);
                bool hasA2 = xMap.TryGetValue(a2, out rawX2) && yMap.TryGetValue(a2, out rawY2);

                if (!hasA1 || !hasA2)
                {
                    continue;
                }

                double sx1 = TransformX(rawX1, minX, scale, offsetX);
                double sy1 = TransformY(rawY1, minY, maxY, scale, offsetY);
                double sx2 = TransformX(rawX2, minX, scale, offsetX);
                double sy2 = TransformY(rawY2, minY, maxY, scale, offsetY);

                int order = 1;
                if (orderArray != null && i < orderArray.Length)
                {
                    order = orderArray[i];
                }

                if (order == 2)
                {
                    double pdx;
                    double pdy;
                    GetPerpendicularOffset(sx1, sy1, sx2, sy2, DoubleOffset, out pdx, out pdy);
                    AppendLine(sb, sx1 + pdx, sy1 + pdy, sx2 + pdx, sy2 + pdy);
                    AppendLine(sb, sx1 - pdx, sy1 - pdy, sx2 - pdx, sy2 - pdy);
                }
                else if (order == 3)
                {
                    double pdx;
                    double pdy;
                    GetPerpendicularOffset(sx1, sy1, sx2, sy2, TripleOffset, out pdx, out pdy);
                    AppendLine(sb, sx1, sy1, sx2, sy2);
                    AppendLine(sb, sx1 + pdx, sy1 + pdy, sx2 + pdx, sy2 + pdy);
                    AppendLine(sb, sx1 - pdx, sy1 - pdy, sx2 - pdx, sy2 - pdy);
                }
                else
                {
                    AppendLine(sb, sx1, sy1, sx2, sy2);
                }
            }
        }

        private static void AppendLine(StringBuilder sb, double x1, double y1, double x2, double y2)
        {
            sb.Append(string.Format(
                CultureInfo.InvariantCulture,
                "  <line x1=\"{0:F2}\" y1=\"{1:F2}\" x2=\"{2:F2}\" y2=\"{3:F2}\" stroke=\"{4}\" stroke-width=\"{5}\" stroke-linecap=\"round\"/>\n",
                x1,
                y1,
                x2,
                y2,
                BondColor,
                BondStrokeWidth));
        }

        private static void AppendAtoms(
            StringBuilder sb,
            PubChemAtoms atoms,
            Dictionary<int, double> xMap,
            Dictionary<int, double> yMap,
            Dictionary<int, int> atomElementMap,
            double minX,
            double minY,
            double maxY,
            double scale,
            double offsetX,
            double offsetY,
            HashSet<int> hiddenHydrogens)
        {
            if (atoms == null || atoms.AtomIds == null)
            {
                return;
            }

            for (int i = 0; i < atoms.AtomIds.Length; i++)
            {
                int atomId = atoms.AtomIds[i];

                if (hiddenHydrogens.Contains(atomId))
                {
                    continue;
                }

                double rawX = 0;
                double rawY = 0;
                bool hasCoords = xMap.TryGetValue(atomId, out rawX) && yMap.TryGetValue(atomId, out rawY);
                if (!hasCoords)
                {
                    continue;
                }

                int element = 6;
                if (atomElementMap.ContainsKey(atomId))
                {
                    element = atomElementMap[atomId];
                }

                if (element == 6)
                {
                    continue;
                }

                double sx = TransformX(rawX, minX, scale, offsetX);
                double sy = TransformY(rawY, minY, maxY, scale, offsetY);
                string symbol = ElementData.GetSymbol(element);
                string hexColor = GetSvgAtomColor(element);

                sb.Append(string.Format(
                    CultureInfo.InvariantCulture,
                    "  <circle cx=\"{0:F2}\" cy=\"{1:F2}\" r=\"{2:F2}\" fill=\"{3}\"/>\n",
                    sx,
                    sy,
                    HeteroBackgroundRadius,
                    BackgroundColor));

                sb.Append(string.Format(
                    CultureInfo.InvariantCulture,
                    "  <text x=\"{0:F2}\" y=\"{1:F2}\" text-anchor=\"middle\" dominant-baseline=\"central\" font-family=\"Arial, sans-serif\" font-size=\"{2}\" font-weight=\"bold\" fill=\"{3}\">{4}</text>\n",
                    sx,
                    sy,
                    FontSize,
                    hexColor,
                    symbol));
            }
        }

        private static readonly Dictionary<int, string> SvgAtomColors = new Dictionary<int, string>
        {
            { 1, "#FFFFFF" },
            { 6, "#AAAAAA" },
            { 7, "#4444FF" },
            { 8, "#FF4444" },
            { 9, "#66DD44" },
            { 15, "#FF9933" },
            { 16, "#FFDD44" },
            { 17, "#44FF44" },
            { 35, "#CC4444" },
            { 53, "#AA44AA" },
        };

        private static string GetSvgAtomColor(int atomicNumber)
        {
            string color;
            if (SvgAtomColors.TryGetValue(atomicNumber, out color))
            {
                return color;
            }

            return "#FF88CC";
        }

        private static void GetPerpendicularOffset(
            double x1,
            double y1,
            double x2,
            double y2,
            double offset,
            out double pdx,
            out double pdy)
        {
            double dx = x2 - x1;
            double dy = y2 - y1;
            double length = Math.Sqrt((dx * dx) + (dy * dy));

            if (length < 0.001)
            {
                pdx = 0;
                pdy = offset;
                return;
            }

            pdx = (-dy / length) * offset;
            pdy = (dx / length) * offset;
        }

        private static double TransformX(double x, double minX, double scale, double offsetX)
        {
            return ((x - minX) * scale) + offsetX;
        }

        private static double TransformY(double y, double minY, double maxY, double scale, double offsetY)
        {
            return ((maxY - y) * scale) + offsetY;
        }
    }
}
