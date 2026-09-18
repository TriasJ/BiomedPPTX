using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class BioArtItem
    {
        public int Id { get; set; }
        public string BioArtId { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string Url { get; set; }
        public string License { get; set; }
        public List<string> Keywords { get; set; }
        public List<BioArtFileEndpoint> FileEndpoints { get; set; }
        public bool HasFiles { get; set; }
    }

    public class BioArtFileEndpoint
    {
        [JsonProperty("file_id")]
        public string FileId { get; set; }

        [JsonProperty("url")]
        public string FileUrl { get; set; }
    }

    public class BioArtFetcher
    {
        private readonly string _indexPath;
        private Dictionary<string, BioArtItem> _index;
        private readonly string _cachePath;

        public BioArtFetcher(string indexPath, string cachePath)
        {
            _indexPath = indexPath;
            _cachePath = cachePath;

            if (!Directory.Exists(_cachePath))
            {
                Directory.CreateDirectory(_cachePath);
            }
        }

        public void LoadIndex()
        {
            if (_index != null) return;

            if (!File.Exists(_indexPath))
            {
                _index = new Dictionary<string, BioArtItem>();
                return;
            }

            string json = File.ReadAllText(_indexPath);
            var raw = JsonConvert.DeserializeObject<Dictionary<string, JObject>>(json);
            _index = new Dictionary<string, BioArtItem>();

            foreach (var kvp in raw)
            {
                var item = new BioArtItem
                {
                    Id = kvp.Value.Value<int>("id"),
                    BioArtId = kvp.Value.Value<string>("bioart_id") ?? "",
                    Title = kvp.Value.Value<string>("title") ?? "",
                    Description = kvp.Value.Value<string>("description") ?? "",
                    Url = kvp.Value.Value<string>("url") ?? "",
                    License = kvp.Value.Value<string>("license") ?? "Unknown",
                    HasFiles = kvp.Value.Value<bool>("has_files"),
                    Keywords = kvp.Value["keywords"]?.ToObject<List<string>>() ?? new List<string>(),
                    FileEndpoints = kvp.Value["file_endpoints"]?.ToObject<List<BioArtFileEndpoint>>() ?? new List<BioArtFileEndpoint>()
                };
                _index[kvp.Key] = item;
            }
        }

        public List<BioArtItem> Search(string query, int limit = 50)
        {
            LoadIndex();

            if (string.IsNullOrWhiteSpace(query) || _index.Count == 0)
            {
                return new List<BioArtItem>();
            }

            string[] terms = query.ToLower().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            var scored = new List<Tuple<BioArtItem, int>>();
            foreach (var item in _index.Values)
            {
                int score = 0;
                string titleLower = item.Title.ToLower();
                string descLower = item.Description.ToLower();

                foreach (string term in terms)
                {
                    if (item.BioArtId.ToLower().Contains(term)) score += 100;
                    if (titleLower.Contains(term)) score += 10;
                    if (descLower.Contains(term)) score += 7;
                    if (item.Keywords.Any(k => k.ToLower() == term)) score += 5;
                }

                if (item.HasFiles) score += 1;

                if (score > 0)
                {
                    scored.Add(Tuple.Create(item, score));
                }
            }

            return scored
                .OrderByDescending(s => s.Item2)
                .Take(limit)
                .Select(s => s.Item1)
                .ToList();
        }

        public string DownloadImage(BioArtItem item)
        {
            if (item.FileEndpoints == null || item.FileEndpoints.Count == 0)
            {
                return null;
            }

            string cachedPath = Path.Combine(_cachePath, $"bioart_{item.Id}.png");
            if (File.Exists(cachedPath))
            {
                return cachedPath;
            }

            foreach (var endpoint in item.FileEndpoints)
            {
                try
                {
                    using (var client = new WebClient())
                    {
                        client.Headers.Add("User-Agent", "BiomedPPTX");
                        byte[] data = client.DownloadData(endpoint.FileUrl);

                        string ext = DetectFileType(data);
                        if (ext == "png" || ext == "jpg" || ext == "svg")
                        {
                            string filePath = Path.Combine(_cachePath, $"bioart_{item.Id}.{ext}");
                            File.WriteAllBytes(filePath, data);
                            return filePath;
                        }
                    }
                }
                catch
                {
                    continue;
                }
            }

            return null;
        }

        public int IndexCount
        {
            get
            {
                LoadIndex();
                return _index.Count;
            }
        }

        private string DetectFileType(byte[] data)
        {
            if (data.Length < 4) return "bin";

            if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
                return "png";
            if (data[0] == 0xFF && data[1] == 0xD8)
                return "jpg";
            if (data[0] == 0x3C) // '<' — likely SVG or XML
            {
                string header = System.Text.Encoding.UTF8.GetString(data, 0, Math.Min(100, data.Length));
                if (header.Contains("<svg") || header.Contains("<?xml"))
                    return "svg";
            }
            if (data[0] == 0x25 && data[1] == 0x50 && data[2] == 0x44 && data[3] == 0x46)
                return "pdf";

            return "png";
        }
    }
}
