using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class BioArtFetcher
    {
        private readonly string _indexPath;
        private Dictionary<string, BioArtItem> _index;
        private readonly string _cachePath;
        private DateTime _lastDownload;
        private bool _isDownloading;
        private const int RateLimitMs = 2000;

        public BioArtFetcher(string indexPath, string cachePath)
        {
            _indexPath = indexPath;
            _cachePath = cachePath;
            _lastDownload = DateTime.MinValue;
            _isDownloading = false;

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                | SecurityProtocolType.Tls11
                | SecurityProtocolType.Tls;

            if (!Directory.Exists(_cachePath))
            {
                Directory.CreateDirectory(_cachePath);
            }
        }

        public void LoadIndex()
        {
            if (_index != null)
            {
                return;
            }

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
                    Keywords = kvp.Value["keywords"] != null ? kvp.Value["keywords"].ToObject<List<string>>() : new List<string>(),
                    FileEndpoints = kvp.Value["file_endpoints"] != null ? kvp.Value["file_endpoints"].ToObject<List<BioArtFileEndpoint>>() : new List<BioArtFileEndpoint>()
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
                    if (item.BioArtId.ToLower().Contains(term))
                    {
                        score += 100;
                    }

                    if (titleLower.Contains(term))
                    {
                        score += 10;
                    }

                    if (descLower.Contains(term))
                    {
                        score += 7;
                    }

                    if (item.Keywords.Any(k => k.ToLower() == term))
                    {
                        score += 5;
                    }
                }

                if (item.HasFiles)
                {
                    score += 1;
                }

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

        public BioArtItem FindById(int id)
        {
            LoadIndex();
            string key = id.ToString();
            BioArtItem result;
            if (_index.TryGetValue(key, out result))
            {
                return result;
            }

            return null;
        }

        public string GetCachedPath(int bioArtId)
        {
            string[] exts = { "png", "svg", "jpg" };
            foreach (string ext in exts)
            {
                string cached = Path.Combine(_cachePath, string.Format("bioart_{0}.{1}", bioArtId, ext));
                if (File.Exists(cached))
                {
                    return cached;
                }
            }

            return null;
        }

        public bool IsDownloading
        {
            get { return _isDownloading; }
        }

        public string DownloadImage(BioArtItem item, string preferredFormat = "svg")
        {
            if (item == null || item.FileEndpoints == null || item.FileEndpoints.Count == 0)
            {
                return null;
            }

            string existing = GetCachedPath(item.Id);
            if (existing != null)
            {
                return existing;
            }

            _isDownloading = true;
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                    | SecurityProtocolType.Tls11
                    | SecurityProtocolType.Tls;

                EnforceRateLimit();

                string[] preferOrder = preferredFormat == "svg"
                    ? new[] { "svg", "png", "jpg" }
                    : new[] { "png", "svg", "jpg" };

                foreach (string targetExt in preferOrder)
                {
                    foreach (var endpoint in item.FileEndpoints)
                    {
                        try
                        {
                            using (var client = new WebClient())
                            {
                                client.Headers.Add("User-Agent", "BiomedPPTX/1.0");
                                byte[] data = client.DownloadData(endpoint.FileUrl);
                                string ext = DetectFileType(data);

                                if (ext == targetExt)
                                {
                                    string filePath = Path.Combine(_cachePath, string.Format("bioart_{0}.{1}", item.Id, ext));
                                    File.WriteAllBytes(filePath, data);
                                    _lastDownload = DateTime.Now;
                                    return filePath;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            continue;
                        }
                    }
                }

                EnforceRateLimit();

                foreach (var endpoint in item.FileEndpoints)
                {
                    try
                    {
                        using (var client = new WebClient())
                        {
                            client.Headers.Add("User-Agent", "BiomedPPTX/1.0");
                            byte[] data = client.DownloadData(endpoint.FileUrl);
                            string ext = DetectFileType(data);
                            if (ext == "png" || ext == "jpg" || ext == "svg")
                            {
                                string filePath = Path.Combine(_cachePath, string.Format("bioart_{0}.{1}", item.Id, ext));
                                File.WriteAllBytes(filePath, data);
                                _lastDownload = DateTime.Now;
                                return filePath;
                            }
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }

                return null;
            }
            finally
            {
                _isDownloading = false;
            }
        }

        public int IndexCount
        {
            get
            {
                LoadIndex();
                return _index.Count;
            }
        }

        private void EnforceRateLimit()
        {
            TimeSpan elapsed = DateTime.Now - _lastDownload;
            if (elapsed.TotalMilliseconds < RateLimitMs)
            {
                int waitMs = RateLimitMs - (int)elapsed.TotalMilliseconds;
                Thread.Sleep(waitMs);
            }
        }

        private string DetectFileType(byte[] data)
        {
            if (data.Length < 4)
            {
                return "bin";
            }

            if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
            {
                return "png";
            }

            if (data[0] == 0xFF && data[1] == 0xD8)
            {
                return "jpg";
            }

            if (data[0] == 0x3C)
            {
                string header = System.Text.Encoding.UTF8.GetString(data, 0, Math.Min(100, data.Length));
                if (header.Contains("<svg") || header.Contains("<?xml"))
                {
                    return "svg";
                }
            }

            if (data[0] == 0x25 && data[1] == 0x50 && data[2] == 0x44 && data[3] == 0x46)
            {
                return "pdf";
            }

            return "png";
        }
    }
}
