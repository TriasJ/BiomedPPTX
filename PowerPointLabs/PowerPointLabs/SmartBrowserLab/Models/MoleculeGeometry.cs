using System;
using System.Collections.Generic;
using System.Windows.Media;
using System.Windows.Media.Media3D;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public static class MoleculeGeometry
    {
        private const double AtomRadiusScale = 0.65;
        private const double HydrogenRadiusScale = 0.35;
        private const double SingleBondRadius = 0.08;
        private const double DoubleBondRadius = 0.06;
        private const double DoubleBondOffset = 0.1;
        private const double TripleBondRadius = 0.05;
        private const double TripleBondOffset = 0.12;
        private const int DefaultSubdivisions = 2;
        private const int DefaultCylinderSides = 12;
        private const double ParallelThreshold = 0.95;

        public static Model3DGroup BuildMoleculeModel(PubChemCompoundData compound)
        {
            Model3DGroup group = new Model3DGroup();

            if (compound == null)
            {
                return group;
            }

            if (compound.Atoms == null || compound.Atoms.AtomIds == null || compound.Atoms.Elements == null)
            {
                return group;
            }

            if (compound.Coords == null || compound.Coords.Length == 0)
            {
                return group;
            }

            PubChemCoords coords = compound.Coords[0];
            if (coords == null || coords.Conformers == null || coords.Conformers.Length == 0)
            {
                return group;
            }

            PubChemConformer conformer = coords.Conformers[0];
            if (conformer == null || conformer.X == null || conformer.Y == null || conformer.Z == null)
            {
                return group;
            }

            if (coords.AtomIds == null)
            {
                return group;
            }

            // Build atom ID to position lookup
            Dictionary<int, Point3D> positionMap = new Dictionary<int, Point3D>();
            for (int i = 0; i < coords.AtomIds.Length; i++)
            {
                if (i < conformer.X.Length && i < conformer.Y.Length && i < conformer.Z.Length)
                {
                    positionMap[coords.AtomIds[i]] = new Point3D(conformer.X[i], conformer.Y[i], conformer.Z[i]);
                }
            }

            // Build atom ID to element lookup
            Dictionary<int, int> elementMap = new Dictionary<int, int>();
            for (int i = 0; i < compound.Atoms.AtomIds.Length; i++)
            {
                if (i < compound.Atoms.Elements.Length)
                {
                    elementMap[compound.Atoms.AtomIds[i]] = compound.Atoms.Elements[i];
                }
            }

            // Create atom spheres
            MeshGeometry3D sphereTemplate = CreateIcosphere(1.0, DefaultSubdivisions);

            foreach (int atomId in compound.Atoms.AtomIds)
            {
                Point3D position;
                if (!positionMap.TryGetValue(atomId, out position))
                {
                    continue;
                }

                int atomicNumber;
                if (!elementMap.TryGetValue(atomId, out atomicNumber))
                {
                    continue;
                }

                double rscale = atomicNumber == 1 ? HydrogenRadiusScale : AtomRadiusScale;
                double radius = ElementData.GetRadius(atomicNumber) * rscale;
                Color color = ElementData.GetColor(atomicNumber);

                MeshGeometry3D scaledSphere = ScaleAndTranslateMesh(sphereTemplate, radius, position.X, position.Y, position.Z);
                scaledSphere.Freeze();

                SolidColorBrush brush = new SolidColorBrush(color);
                brush.Freeze();
                DiffuseMaterial mat = new DiffuseMaterial(brush);
                mat.Freeze();

                GeometryModel3D atomModel = new GeometryModel3D();
                atomModel.Geometry = scaledSphere;
                atomModel.Material = mat;
                atomModel.BackMaterial = mat;
                atomModel.Freeze();

                group.Children.Add(atomModel);
            }

            // Create bonds
            if (compound.Bonds != null && compound.Bonds.AtomId1 != null && compound.Bonds.AtomId2 != null)
            {
                int bondCount = compound.Bonds.AtomId1.Length;
                Color bondColor = Color.FromRgb(160, 160, 160);
                SolidColorBrush bondBrush = new SolidColorBrush(bondColor);
                bondBrush.Freeze();
                DiffuseMaterial bondMaterial = new DiffuseMaterial(bondBrush);
                bondMaterial.Freeze();

                for (int i = 0; i < bondCount; i++)
                {
                    if (i >= compound.Bonds.AtomId2.Length)
                    {
                        break;
                    }

                    int aid1 = compound.Bonds.AtomId1[i];
                    int aid2 = compound.Bonds.AtomId2[i];

                    Point3D start;
                    Point3D end;
                    if (!positionMap.TryGetValue(aid1, out start) || !positionMap.TryGetValue(aid2, out end))
                    {
                        continue;
                    }

                    int order = 1;
                    if (compound.Bonds.Order != null && i < compound.Bonds.Order.Length)
                    {
                        order = compound.Bonds.Order[i];
                    }

                    if (order == 1)
                    {
                        MeshGeometry3D cylinder = CreateCylinder(start, end, SingleBondRadius, DefaultCylinderSides);
                        cylinder.Freeze();
                        GeometryModel3D bondModel = new GeometryModel3D();
                        bondModel.Geometry = cylinder;
                        bondModel.Material = bondMaterial;
                        bondModel.BackMaterial = bondMaterial;
                        bondModel.Freeze();
                        group.Children.Add(bondModel);
                    }
                    else if (order == 2)
                    {
                        Vector3D bondDir = end - start;
                        Vector3D perp = GetPerpendicularVector(bondDir);
                        perp.Normalize();

                        Vector3D offset = perp * DoubleBondOffset;

                        Point3D start1 = new Point3D(start.X + offset.X, start.Y + offset.Y, start.Z + offset.Z);
                        Point3D end1 = new Point3D(end.X + offset.X, end.Y + offset.Y, end.Z + offset.Z);
                        Point3D start2 = new Point3D(start.X - offset.X, start.Y - offset.Y, start.Z - offset.Z);
                        Point3D end2 = new Point3D(end.X - offset.X, end.Y - offset.Y, end.Z - offset.Z);

                        MeshGeometry3D cyl1 = CreateCylinder(start1, end1, DoubleBondRadius, DefaultCylinderSides);
                        cyl1.Freeze();
                        MeshGeometry3D cyl2 = CreateCylinder(start2, end2, DoubleBondRadius, DefaultCylinderSides);
                        cyl2.Freeze();

                        GeometryModel3D model1 = new GeometryModel3D();
                        model1.Geometry = cyl1;
                        model1.Material = bondMaterial;
                        model1.BackMaterial = bondMaterial;
                        model1.Freeze();
                        group.Children.Add(model1);

                        GeometryModel3D model2 = new GeometryModel3D();
                        model2.Geometry = cyl2;
                        model2.Material = bondMaterial;
                        model2.BackMaterial = bondMaterial;
                        model2.Freeze();
                        group.Children.Add(model2);
                    }
                    else if (order == 3)
                    {
                        Vector3D bondDir = end - start;
                        Vector3D perp = GetPerpendicularVector(bondDir);
                        perp.Normalize();

                        Vector3D offset = perp * TripleBondOffset;

                        MeshGeometry3D cylCenter = CreateCylinder(start, end, TripleBondRadius, DefaultCylinderSides);
                        cylCenter.Freeze();
                        GeometryModel3D modelCenter = new GeometryModel3D();
                        modelCenter.Geometry = cylCenter;
                        modelCenter.Material = bondMaterial;
                        modelCenter.BackMaterial = bondMaterial;
                        modelCenter.Freeze();
                        group.Children.Add(modelCenter);

                        Point3D start1 = new Point3D(start.X + offset.X, start.Y + offset.Y, start.Z + offset.Z);
                        Point3D end1 = new Point3D(end.X + offset.X, end.Y + offset.Y, end.Z + offset.Z);
                        Point3D start2 = new Point3D(start.X - offset.X, start.Y - offset.Y, start.Z - offset.Z);
                        Point3D end2 = new Point3D(end.X - offset.X, end.Y - offset.Y, end.Z - offset.Z);

                        MeshGeometry3D cyl1 = CreateCylinder(start1, end1, TripleBondRadius, DefaultCylinderSides);
                        cyl1.Freeze();
                        MeshGeometry3D cyl2 = CreateCylinder(start2, end2, TripleBondRadius, DefaultCylinderSides);
                        cyl2.Freeze();

                        GeometryModel3D model1 = new GeometryModel3D();
                        model1.Geometry = cyl1;
                        model1.Material = bondMaterial;
                        model1.BackMaterial = bondMaterial;
                        model1.Freeze();
                        group.Children.Add(model1);

                        GeometryModel3D model2 = new GeometryModel3D();
                        model2.Geometry = cyl2;
                        model2.Material = bondMaterial;
                        model2.BackMaterial = bondMaterial;
                        model2.Freeze();
                        group.Children.Add(model2);
                    }
                    else
                    {
                        MeshGeometry3D cylinder = CreateCylinder(start, end, SingleBondRadius, DefaultCylinderSides);
                        cylinder.Freeze();
                        GeometryModel3D bondModel = new GeometryModel3D();
                        bondModel.Geometry = cylinder;
                        bondModel.Material = bondMaterial;
                        bondModel.BackMaterial = bondMaterial;
                        bondModel.Freeze();
                        group.Children.Add(bondModel);
                    }
                }
            }

            group.Freeze();
            return group;
        }

        private static MeshGeometry3D CreateIcosphere(double radius, int subdivisions)
        {
            List<Point3D> vertices = new List<Point3D>();
            List<int> indices = new List<int>();

            // Golden ratio
            double phi = (1.0 + Math.Sqrt(5.0)) / 2.0;

            // 12 vertices of a regular icosahedron, normalized to unit sphere
            AddNormalizedVertex(vertices, -1.0, phi, 0.0);
            AddNormalizedVertex(vertices, 1.0, phi, 0.0);
            AddNormalizedVertex(vertices, -1.0, -phi, 0.0);
            AddNormalizedVertex(vertices, 1.0, -phi, 0.0);

            AddNormalizedVertex(vertices, 0.0, -1.0, phi);
            AddNormalizedVertex(vertices, 0.0, 1.0, phi);
            AddNormalizedVertex(vertices, 0.0, -1.0, -phi);
            AddNormalizedVertex(vertices, 0.0, 1.0, -phi);

            AddNormalizedVertex(vertices, phi, 0.0, -1.0);
            AddNormalizedVertex(vertices, phi, 0.0, 1.0);
            AddNormalizedVertex(vertices, -phi, 0.0, -1.0);
            AddNormalizedVertex(vertices, -phi, 0.0, 1.0);

            // 20 triangular faces of the icosahedron
            // Faces around vertex 0
            indices.AddRange(new int[] { 0, 11, 5 });
            indices.AddRange(new int[] { 0, 5, 1 });
            indices.AddRange(new int[] { 0, 1, 7 });
            indices.AddRange(new int[] { 0, 7, 10 });
            indices.AddRange(new int[] { 0, 10, 11 });

            // Adjacent faces
            indices.AddRange(new int[] { 1, 5, 9 });
            indices.AddRange(new int[] { 5, 11, 4 });
            indices.AddRange(new int[] { 11, 10, 2 });
            indices.AddRange(new int[] { 10, 7, 6 });
            indices.AddRange(new int[] { 7, 1, 8 });

            // Faces around vertex 3
            indices.AddRange(new int[] { 3, 9, 4 });
            indices.AddRange(new int[] { 3, 4, 2 });
            indices.AddRange(new int[] { 3, 2, 6 });
            indices.AddRange(new int[] { 3, 6, 8 });
            indices.AddRange(new int[] { 3, 8, 9 });

            // Remaining faces
            indices.AddRange(new int[] { 4, 9, 5 });
            indices.AddRange(new int[] { 2, 4, 11 });
            indices.AddRange(new int[] { 6, 2, 10 });
            indices.AddRange(new int[] { 8, 6, 7 });
            indices.AddRange(new int[] { 9, 8, 1 });

            // Subdivide
            Dictionary<long, int> midpointCache = new Dictionary<long, int>();

            for (int s = 0; s < subdivisions; s++)
            {
                List<int> newIndices = new List<int>();
                midpointCache.Clear();

                for (int i = 0; i < indices.Count; i += 3)
                {
                    int v0 = indices[i];
                    int v1 = indices[i + 1];
                    int v2 = indices[i + 2];

                    int a = GetMidpointIndex(vertices, midpointCache, v0, v1);
                    int b = GetMidpointIndex(vertices, midpointCache, v1, v2);
                    int c = GetMidpointIndex(vertices, midpointCache, v2, v0);

                    newIndices.AddRange(new int[] { v0, a, c });
                    newIndices.AddRange(new int[] { v1, b, a });
                    newIndices.AddRange(new int[] { v2, c, b });
                    newIndices.AddRange(new int[] { a, b, c });
                }

                indices = newIndices;
            }

            // Build mesh
            MeshGeometry3D mesh = new MeshGeometry3D();

            for (int i = 0; i < vertices.Count; i++)
            {
                Point3D v = vertices[i];
                mesh.Positions.Add(new Point3D(v.X * radius, v.Y * radius, v.Z * radius));
                mesh.Normals.Add(new Vector3D(v.X, v.Y, v.Z));
            }

            for (int i = 0; i < indices.Count; i++)
            {
                mesh.TriangleIndices.Add(indices[i]);
            }

            return mesh;
        }

        private static MeshGeometry3D CreateCylinder(Point3D start, Point3D end, double radius, int sides)
        {
            MeshGeometry3D mesh = new MeshGeometry3D();

            Vector3D direction = end - start;
            double length = direction.Length;

            if (length < 1e-10)
            {
                return mesh;
            }

            direction.Normalize();

            Vector3D perpA = GetPerpendicularVector(direction);
            perpA.Normalize();

            Vector3D perpB = Vector3D.CrossProduct(direction, perpA);
            perpB.Normalize();

            // Generate circle vertices at start and end
            double angleStep = 2.0 * Math.PI / sides;

            for (int i = 0; i < sides; i++)
            {
                double angle = i * angleStep;
                double cosA = Math.Cos(angle);
                double sinA = Math.Sin(angle);

                // Radial direction for this segment
                Vector3D radial = (perpA * cosA) + (perpB * sinA);

                // Bottom circle vertex (at start)
                Point3D bottomVertex = new Point3D(
                    start.X + (radial.X * radius),
                    start.Y + (radial.Y * radius),
                    start.Z + (radial.Z * radius));

                // Top circle vertex (at end)
                Point3D topVertex = new Point3D(
                    end.X + (radial.X * radius),
                    end.Y + (radial.Y * radius),
                    end.Z + (radial.Z * radius));

                mesh.Positions.Add(bottomVertex);
                mesh.Normals.Add(radial);

                mesh.Positions.Add(topVertex);
                mesh.Normals.Add(radial);
            }

            // Side faces: triangle strip connecting bottom and top circles
            for (int i = 0; i < sides; i++)
            {
                int next = (i + 1) % sides;

                int bottomCurrent = i * 2;
                int topCurrent = (i * 2) + 1;
                int bottomNext = next * 2;
                int topNext = (next * 2) + 1;

                // Triangle 1
                mesh.TriangleIndices.Add(bottomCurrent);
                mesh.TriangleIndices.Add(topCurrent);
                mesh.TriangleIndices.Add(bottomNext);

                // Triangle 2
                mesh.TriangleIndices.Add(bottomNext);
                mesh.TriangleIndices.Add(topCurrent);
                mesh.TriangleIndices.Add(topNext);
            }

            // Bottom cap center
            int bottomCenterIndex = mesh.Positions.Count;
            mesh.Positions.Add(start);
            mesh.Normals.Add(-direction);

            // Bottom cap triangles (reversed winding for outward-facing normal)
            for (int i = 0; i < sides; i++)
            {
                int next = (i + 1) % sides;
                int currentBottom = i * 2;
                int nextBottom = next * 2;

                // Add duplicate vertices with cap normal for proper shading
                int capIdx1 = mesh.Positions.Count;
                mesh.Positions.Add(mesh.Positions[currentBottom]);
                mesh.Normals.Add(-direction);

                int capIdx2 = mesh.Positions.Count;
                mesh.Positions.Add(mesh.Positions[nextBottom]);
                mesh.Normals.Add(-direction);

                mesh.TriangleIndices.Add(bottomCenterIndex);
                mesh.TriangleIndices.Add(capIdx2);
                mesh.TriangleIndices.Add(capIdx1);
            }

            // Top cap center
            int topCenterIndex = mesh.Positions.Count;
            mesh.Positions.Add(end);
            mesh.Normals.Add(direction);

            // Top cap triangles
            for (int i = 0; i < sides; i++)
            {
                int next = (i + 1) % sides;
                int currentTop = (i * 2) + 1;
                int nextTop = (next * 2) + 1;

                // Add duplicate vertices with cap normal for proper shading
                int capIdx1 = mesh.Positions.Count;
                mesh.Positions.Add(mesh.Positions[currentTop]);
                mesh.Normals.Add(direction);

                int capIdx2 = mesh.Positions.Count;
                mesh.Positions.Add(mesh.Positions[nextTop]);
                mesh.Normals.Add(direction);

                mesh.TriangleIndices.Add(topCenterIndex);
                mesh.TriangleIndices.Add(capIdx1);
                mesh.TriangleIndices.Add(capIdx2);
            }

            return mesh;
        }

        private static MeshGeometry3D TranslateMesh(MeshGeometry3D mesh, double x, double y, double z)
        {
            MeshGeometry3D result = new MeshGeometry3D();

            for (int i = 0; i < mesh.Positions.Count; i++)
            {
                Point3D p = mesh.Positions[i];
                result.Positions.Add(new Point3D(p.X + x, p.Y + y, p.Z + z));
            }

            for (int i = 0; i < mesh.Normals.Count; i++)
            {
                result.Normals.Add(mesh.Normals[i]);
            }

            for (int i = 0; i < mesh.TriangleIndices.Count; i++)
            {
                result.TriangleIndices.Add(mesh.TriangleIndices[i]);
            }

            return result;
        }

        private static MeshGeometry3D ScaleAndTranslateMesh(MeshGeometry3D mesh, double scale, double x, double y, double z)
        {
            MeshGeometry3D result = new MeshGeometry3D();

            for (int i = 0; i < mesh.Positions.Count; i++)
            {
                Point3D p = mesh.Positions[i];
                result.Positions.Add(new Point3D((p.X * scale) + x, (p.Y * scale) + y, (p.Z * scale) + z));
            }

            for (int i = 0; i < mesh.Normals.Count; i++)
            {
                result.Normals.Add(mesh.Normals[i]);
            }

            for (int i = 0; i < mesh.TriangleIndices.Count; i++)
            {
                result.TriangleIndices.Add(mesh.TriangleIndices[i]);
            }

            return result;
        }

        private static Vector3D GetPerpendicularVector(Vector3D direction)
        {
            Vector3D up = new Vector3D(0, 1, 0);
            Vector3D normalized = direction;
            normalized.Normalize();

            double dot = Math.Abs(Vector3D.DotProduct(normalized, up));

            if (dot > ParallelThreshold)
            {
                up = new Vector3D(1, 0, 0);
            }

            Vector3D perp = Vector3D.CrossProduct(direction, up);
            perp.Normalize();
            return perp;
        }

        private static void AddNormalizedVertex(List<Point3D> vertices, double x, double y, double z)
        {
            double length = Math.Sqrt((x * x) + (y * y) + (z * z));
            vertices.Add(new Point3D(x / length, y / length, z / length));
        }

        private static int GetMidpointIndex(List<Point3D> vertices, Dictionary<long, int> cache, int i1, int i2)
        {
            // Ensure consistent key regardless of order
            int smaller = Math.Min(i1, i2);
            int larger = Math.Max(i1, i2);
            long key = ((long)smaller << 32) + larger;

            int midpointIndex;
            if (cache.TryGetValue(key, out midpointIndex))
            {
                return midpointIndex;
            }

            Point3D p1 = vertices[i1];
            Point3D p2 = vertices[i2];

            double mx = (p1.X + p2.X) / 2.0;
            double my = (p1.Y + p2.Y) / 2.0;
            double mz = (p1.Z + p2.Z) / 2.0;

            // Normalize to unit sphere
            double length = Math.Sqrt((mx * mx) + (my * my) + (mz * mz));
            mx /= length;
            my /= length;
            mz /= length;

            midpointIndex = vertices.Count;
            vertices.Add(new Point3D(mx, my, mz));
            cache[key] = midpointIndex;

            return midpointIndex;
        }
    }
}
