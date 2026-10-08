using System;
using System.Collections.Generic;
using System.Numerics;

using SharpGLTF.Geometry;
using SharpGLTF.Geometry.VertexTypes;
using SharpGLTF.Materials;
using SharpGLTF.Scenes;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public static class MoleculeGlbExporter
    {
        private const double AtomRadiusScale = 0.65;
        private const double HydrogenRadiusScale = 0.35;
        private const float SingleBondRadius = 0.08f;
        private const float DoubleBondRadius = 0.06f;
        private const float DoubleBondOffset = 0.1f;
        private const float TripleBondRadius = 0.05f;
        private const float TripleBondOffset = 0.12f;
        private const int IcoSubdivisions = 2;
        private const int CylinderSides = 12;

        public static void ExportGlb(PubChemCompoundData compound, string outputPath)
        {
            if (compound == null)
            {
                throw new ArgumentNullException("compound");
            }

            if (string.IsNullOrEmpty(outputPath))
            {
                throw new ArgumentNullException("outputPath");
            }

            if (compound.Atoms == null || compound.Atoms.AtomIds == null ||
                compound.Atoms.Elements == null || compound.Coords == null ||
                compound.Coords.Length == 0)
            {
                SaveEmptyGlb(outputPath);
                return;
            }

            PubChemCoords coords = compound.Coords[0];
            if (coords == null || coords.Conformers == null || coords.Conformers.Length == 0 ||
                coords.AtomIds == null)
            {
                SaveEmptyGlb(outputPath);
                return;
            }

            PubChemConformer conformer = coords.Conformers[0];
            if (conformer == null || conformer.X == null || conformer.Y == null || conformer.Z == null)
            {
                SaveEmptyGlb(outputPath);
                return;
            }

            Dictionary<int, Vector3> positionMap = new Dictionary<int, Vector3>();
            for (int i = 0; i < coords.AtomIds.Length; i++)
            {
                if (i < conformer.X.Length && i < conformer.Y.Length && i < conformer.Z.Length)
                {
                    positionMap[coords.AtomIds[i]] = new Vector3(
                        (float)conformer.X[i],
                        (float)conformer.Y[i],
                        (float)conformer.Z[i]);
                }
            }

            Dictionary<int, int> elementMap = new Dictionary<int, int>();
            for (int i = 0; i < compound.Atoms.AtomIds.Length; i++)
            {
                if (i < compound.Atoms.Elements.Length)
                {
                    elementMap[compound.Atoms.AtomIds[i]] = compound.Atoms.Elements[i];
                }
            }

            List<Vector3> sphereVerts;
            List<Vector3> sphereNormals;
            List<int> sphereIndices;
            GenerateIcosphere(1.0f, IcoSubdivisions, out sphereVerts, out sphereNormals, out sphereIndices);

            var mesh = new MeshBuilder<VertexPositionNormal>("molecule");

            Dictionary<int, MaterialBuilder> materialCache = new Dictionary<int, MaterialBuilder>();

            foreach (int atomId in compound.Atoms.AtomIds)
            {
                Vector3 position;
                if (!positionMap.TryGetValue(atomId, out position))
                {
                    continue;
                }

                int atomicNumber;
                if (!elementMap.TryGetValue(atomId, out atomicNumber))
                {
                    continue;
                }

                double scale = atomicNumber == 1 ? HydrogenRadiusScale : AtomRadiusScale;
                float radius = (float)(ElementData.GetRadius(atomicNumber) * scale);

                MaterialBuilder material;
                if (!materialCache.TryGetValue(atomicNumber, out material))
                {
                    System.Windows.Media.Color wpfColor = ElementData.GetColor(atomicNumber);
                    Vector4 color = new Vector4(wpfColor.R / 255f, wpfColor.G / 255f, wpfColor.B / 255f, 1f);
                    material = new MaterialBuilder(string.Format("element_{0}", atomicNumber))
                        .WithMetallicRoughnessShader()
                        .WithBaseColor(color)
                        .WithMetallicRoughness(0.05f, 0.55f);
                    materialCache[atomicNumber] = material;
                }

                var prim = mesh.UsePrimitive(material);

                for (int i = 0; i < sphereIndices.Count; i += 3)
                {
                    Vector3 p0 = sphereVerts[sphereIndices[i]] * radius + position;
                    Vector3 p1 = sphereVerts[sphereIndices[i + 1]] * radius + position;
                    Vector3 p2 = sphereVerts[sphereIndices[i + 2]] * radius + position;

                    Vector3 n0 = sphereNormals[sphereIndices[i]];
                    Vector3 n1 = sphereNormals[sphereIndices[i + 1]];
                    Vector3 n2 = sphereNormals[sphereIndices[i + 2]];

                    prim.AddTriangle(
                        new VertexPositionNormal(p0, n0),
                        new VertexPositionNormal(p1, n1),
                        new VertexPositionNormal(p2, n2));
                }
            }

            if (compound.Bonds != null && compound.Bonds.AtomId1 != null &&
                compound.Bonds.AtomId2 != null)
            {
                Vector4 bondColor = new Vector4(0.63f, 0.63f, 0.63f, 1f);
                MaterialBuilder bondMaterial = new MaterialBuilder("bonds")
                    .WithMetallicRoughnessShader()
                    .WithBaseColor(bondColor)
                    .WithMetallicRoughness(0.1f, 0.6f);

                var bondPrim = mesh.UsePrimitive(bondMaterial);

                int bondCount = compound.Bonds.AtomId1.Length;
                for (int i = 0; i < bondCount; i++)
                {
                    if (i >= compound.Bonds.AtomId2.Length)
                    {
                        break;
                    }

                    int aid1 = compound.Bonds.AtomId1[i];
                    int aid2 = compound.Bonds.AtomId2[i];

                    Vector3 start;
                    Vector3 end;
                    if (!positionMap.TryGetValue(aid1, out start) || !positionMap.TryGetValue(aid2, out end))
                    {
                        continue;
                    }

                    int order = 1;
                    if (compound.Bonds.Order != null && i < compound.Bonds.Order.Length)
                    {
                        order = compound.Bonds.Order[i];
                    }

                    if (order == 1 || order < 1 || order > 3)
                    {
                        AddCylinderTriangles(bondPrim, start, end, SingleBondRadius);
                    }
                    else if (order == 2)
                    {
                        Vector3 perp = GetPerpendicular(end - start);
                        Vector3 off = perp * DoubleBondOffset;
                        AddCylinderTriangles(bondPrim, start + off, end + off, DoubleBondRadius);
                        AddCylinderTriangles(bondPrim, start - off, end - off, DoubleBondRadius);
                    }
                    else if (order == 3)
                    {
                        Vector3 perp = GetPerpendicular(end - start);
                        Vector3 off = perp * TripleBondOffset;
                        AddCylinderTriangles(bondPrim, start, end, TripleBondRadius);
                        AddCylinderTriangles(bondPrim, start + off, end + off, TripleBondRadius);
                        AddCylinderTriangles(bondPrim, start - off, end - off, TripleBondRadius);
                    }
                }
            }

            var scene = new SceneBuilder();
            scene.AddRigidMesh(mesh, Matrix4x4.Identity);

            var model = scene.ToGltf2();
            model.SaveGLB(outputPath);
        }

        private static void SaveEmptyGlb(string outputPath)
        {
            var scene = new SceneBuilder();
            var model = scene.ToGltf2();
            model.SaveGLB(outputPath);
        }

        private static void AddCylinderTriangles(
            PrimitiveBuilder<MaterialBuilder, VertexPositionNormal, VertexEmpty, VertexEmpty> prim,
            Vector3 start, Vector3 end, float radius)
        {
            List<Vector3> verts;
            List<Vector3> normals;
            List<int> indices;
            GenerateCylinder(start, end, radius, CylinderSides, out verts, out normals, out indices);

            for (int i = 0; i < indices.Count; i += 3)
            {
                prim.AddTriangle(
                    new VertexPositionNormal(verts[indices[i]], normals[indices[i]]),
                    new VertexPositionNormal(verts[indices[i + 1]], normals[indices[i + 1]]),
                    new VertexPositionNormal(verts[indices[i + 2]], normals[indices[i + 2]]));
            }
        }

        private static void GenerateIcosphere(
            float radius, int subdivisions,
            out List<Vector3> vertices, out List<Vector3> normals, out List<int> indices)
        {
            vertices = new List<Vector3>();
            indices = new List<int>();

            float phi = (1f + (float)Math.Sqrt(5f)) / 2f;

            AddNormVertex(vertices, new Vector3(-1, phi, 0));
            AddNormVertex(vertices, new Vector3(1, phi, 0));
            AddNormVertex(vertices, new Vector3(-1, -phi, 0));
            AddNormVertex(vertices, new Vector3(1, -phi, 0));
            AddNormVertex(vertices, new Vector3(0, -1, phi));
            AddNormVertex(vertices, new Vector3(0, 1, phi));
            AddNormVertex(vertices, new Vector3(0, -1, -phi));
            AddNormVertex(vertices, new Vector3(0, 1, -phi));
            AddNormVertex(vertices, new Vector3(phi, 0, -1));
            AddNormVertex(vertices, new Vector3(phi, 0, 1));
            AddNormVertex(vertices, new Vector3(-phi, 0, -1));
            AddNormVertex(vertices, new Vector3(-phi, 0, 1));

            int[] faces = new int[]
            {
                0, 11, 5, 0, 5, 1, 0, 1, 7, 0, 7, 10, 0, 10, 11,
                1, 5, 9, 5, 11, 4, 11, 10, 2, 10, 7, 6, 7, 1, 8,
                3, 9, 4, 3, 4, 2, 3, 2, 6, 3, 6, 8, 3, 8, 9,
                4, 9, 5, 2, 4, 11, 6, 2, 10, 8, 6, 7, 9, 8, 1
            };

            indices.AddRange(faces);

            Dictionary<long, int> midpointCache = new Dictionary<long, int>();
            for (int s = 0; s < subdivisions; s++)
            {
                List<int> newIndices = new List<int>();
                for (int i = 0; i < indices.Count; i += 3)
                {
                    int v1 = indices[i];
                    int v2 = indices[i + 1];
                    int v3 = indices[i + 2];

                    int a = GetMidpoint(vertices, midpointCache, v1, v2);
                    int b = GetMidpoint(vertices, midpointCache, v2, v3);
                    int c = GetMidpoint(vertices, midpointCache, v3, v1);

                    newIndices.Add(v1);
                    newIndices.Add(a);
                    newIndices.Add(c);
                    newIndices.Add(v2);
                    newIndices.Add(b);
                    newIndices.Add(a);
                    newIndices.Add(v3);
                    newIndices.Add(c);
                    newIndices.Add(b);
                    newIndices.Add(a);
                    newIndices.Add(b);
                    newIndices.Add(c);
                }

                indices = newIndices;
            }

            normals = new List<Vector3>();
            for (int i = 0; i < vertices.Count; i++)
            {
                normals.Add(Vector3.Normalize(vertices[i]));
                vertices[i] = vertices[i] * radius;
            }
        }

        private static void AddNormVertex(List<Vector3> vertices, Vector3 v)
        {
            vertices.Add(Vector3.Normalize(v));
        }

        private static int GetMidpoint(
            List<Vector3> vertices, Dictionary<long, int> cache, int i1, int i2)
        {
            long smallerIndex = Math.Min(i1, i2);
            long greaterIndex = Math.Max(i1, i2);
            long key = (smallerIndex << 20) + greaterIndex;

            int ret;
            if (cache.TryGetValue(key, out ret))
            {
                return ret;
            }

            Vector3 mid = Vector3.Normalize((vertices[i1] + vertices[i2]) * 0.5f);
            int idx = vertices.Count;
            vertices.Add(mid);
            cache[key] = idx;
            return idx;
        }

        private static void GenerateCylinder(
            Vector3 start, Vector3 end, float radius, int sides,
            out List<Vector3> vertices, out List<Vector3> normals, out List<int> indices)
        {
            vertices = new List<Vector3>();
            normals = new List<Vector3>();
            indices = new List<int>();

            Vector3 dir = end - start;
            float length = dir.Length();
            if (length < 0.0001f)
            {
                return;
            }

            dir = Vector3.Normalize(dir);
            Vector3 perp = GetPerpendicular(dir);
            Vector3 biperp = Vector3.Cross(dir, perp);

            for (int i = 0; i < sides; i++)
            {
                float angle = (float)(2.0 * Math.PI * i / sides);
                float cos = (float)Math.Cos(angle);
                float sin = (float)Math.Sin(angle);

                Vector3 normal = perp * cos + biperp * sin;
                Vector3 offset = normal * radius;

                vertices.Add(start + offset);
                normals.Add(normal);
                vertices.Add(end + offset);
                normals.Add(normal);
            }

            for (int i = 0; i < sides; i++)
            {
                int next = (i + 1) % sides;
                int bl = i * 2;
                int tl = i * 2 + 1;
                int br = next * 2;
                int tr = next * 2 + 1;

                indices.Add(bl);
                indices.Add(br);
                indices.Add(tl);
                indices.Add(tl);
                indices.Add(br);
                indices.Add(tr);
            }

            int bottomCenter = vertices.Count;
            vertices.Add(start);
            normals.Add(-dir);
            int topCenter = vertices.Count;
            vertices.Add(end);
            normals.Add(dir);

            for (int i = 0; i < sides; i++)
            {
                int next = (i + 1) % sides;

                int bci = vertices.Count;
                vertices.Add(vertices[i * 2]);
                normals.Add(-dir);
                int bni = vertices.Count;
                vertices.Add(vertices[next * 2]);
                normals.Add(-dir);
                indices.Add(bottomCenter);
                indices.Add(bni);
                indices.Add(bci);

                int tci = vertices.Count;
                vertices.Add(vertices[i * 2 + 1]);
                normals.Add(dir);
                int tni = vertices.Count;
                vertices.Add(vertices[next * 2 + 1]);
                normals.Add(dir);
                indices.Add(topCenter);
                indices.Add(tci);
                indices.Add(tni);
            }
        }

        private static Vector3 GetPerpendicular(Vector3 dir)
        {
            Vector3 up = Math.Abs(Vector3.Dot(dir, Vector3.UnitY)) > 0.95f
                ? Vector3.UnitX
                : Vector3.UnitY;
            Vector3 perp = Vector3.Cross(dir, up);
            return Vector3.Normalize(perp);
        }

        private static IPrimitiveBuilder UsePrimitive(MeshBuilder<VertexPositionNormal> mesh, MaterialBuilder material)
        {
            return mesh.UsePrimitive(material);
        }
    }
}
