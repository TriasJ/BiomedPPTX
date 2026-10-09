using System.IO;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class MoleculeCache
    {
        private readonly string _cachePath;

        public MoleculeCache(string cachePath)
        {
            _cachePath = cachePath;

            if (!Directory.Exists(_cachePath))
            {
                Directory.CreateDirectory(_cachePath);
            }
        }

        public string GetCached3DJson(int cid)
        {
            string path = Path.Combine(_cachePath, string.Format("{0}_3d.json", cid));
            if (File.Exists(path))
            {
                return File.ReadAllText(path);
            }

            return null;
        }

        public void Save3DJson(int cid, string json)
        {
            string path = Path.Combine(_cachePath, string.Format("{0}_3d.json", cid));
            File.WriteAllText(path, json);
        }

        public string GetCached2DJson(int cid)
        {
            string path = Path.Combine(_cachePath, string.Format("{0}_2d.json", cid));
            if (File.Exists(path))
            {
                return File.ReadAllText(path);
            }

            return null;
        }

        public void Save2DJson(int cid, string json)
        {
            string path = Path.Combine(_cachePath, string.Format("{0}_2d.json", cid));
            File.WriteAllText(path, json);
        }

        public string GetCached2DImage(int cid)
        {
            string path = Path.Combine(_cachePath, string.Format("{0}.png", cid));
            if (File.Exists(path))
            {
                return path;
            }

            return null;
        }

        public void Save2DImage(int cid, byte[] png)
        {
            string path = Path.Combine(_cachePath, string.Format("{0}.png", cid));
            File.WriteAllBytes(path, png);
        }

        public string GetCachedHiRes2DImage(int cid)
        {
            string path = Path.Combine(_cachePath, string.Format("{0}_hires.png", cid));
            if (File.Exists(path))
            {
                return path;
            }

            return null;
        }

        public void SaveHiRes2DImage(int cid, byte[] png)
        {
            string path = Path.Combine(_cachePath, string.Format("{0}_hires.png", cid));
            File.WriteAllBytes(path, png);
        }

        public string GetHiRes2DImagePath(int cid)
        {
            return Path.Combine(_cachePath, string.Format("{0}_hires.png", cid));
        }

        public string GetCachedGlb(int cid)
        {
            string path = Path.Combine(_cachePath, string.Format("{0}.glb", cid));
            if (File.Exists(path))
            {
                return path;
            }

            return null;
        }

        public string GetGlbPath(int cid)
        {
            return Path.Combine(_cachePath, string.Format("{0}.glb", cid));
        }

        public string GetCachedSvg(int cid)
        {
            string path = Path.Combine(_cachePath, string.Format("{0}.svg", cid));
            if (File.Exists(path))
            {
                return path;
            }

            return null;
        }

        public string GetSvgPath(int cid)
        {
            return Path.Combine(_cachePath, string.Format("{0}.svg", cid));
        }
    }
}
