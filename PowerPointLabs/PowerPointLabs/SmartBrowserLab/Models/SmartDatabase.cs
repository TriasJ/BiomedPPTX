using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;

using Newtonsoft.Json;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class SmartDatabase : IDisposable
    {
        private SQLiteConnection _conn;
        private readonly string _basePath;
        private bool _hasFts5;

        public SmartDatabase(string dbPath, string basePath)
        {
            _basePath = basePath;
            _conn = new SQLiteConnection(string.Format("Data Source={0};Version=3;Read Only=True;", dbPath));
            _conn.Open();
            _hasFts5 = CheckFts5Support();
        }

        public List<IllustrationItem> Search(string query, int limit = 100)
        {
            if (_hasFts5)
            {
                return SearchFts5(query, limit);
            }

            return SearchLike(query, limit);
        }

        public List<IllustrationItem> GetByTopic(int topicId, int offset = 0, int limit = 100)
        {
            const string sql = @"
                SELECT i.id, i.name, i.description, i.pptx_file, i.pptx_slide, i.pptx_shape_index,
                       i.svg_path, i.png_path, i.width, i.height,
                       tp.name as topic, sl.title as slide_title
                FROM illustrations i
                JOIN slides sl ON i.slide_id = sl.id
                JOIN topics tp ON sl.topic_id = tp.id
                WHERE tp.id = @topicId
                ORDER BY i.name
                LIMIT @limit OFFSET @offset";

            var results = new List<IllustrationItem>();
            using (var cmd = new SQLiteCommand(sql, _conn))
            {
                cmd.Parameters.AddWithValue("@topicId", topicId);
                cmd.Parameters.AddWithValue("@limit", limit);
                cmd.Parameters.AddWithValue("@offset", offset);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        results.Add(ReadIllustration(reader));
                    }
                }
            }
            return results;
        }

        public List<IllustrationItem> GetByTag(string tagName, int offset = 0, int limit = 100)
        {
            const string sql = @"
                SELECT i.id, i.name, i.description, i.pptx_file, i.pptx_slide, i.pptx_shape_index,
                       i.svg_path, i.png_path, i.width, i.height,
                       tp.name as topic, sl.title as slide_title
                FROM illustrations i
                JOIN illustration_tags it ON it.illustration_id = i.id
                JOIN tags t ON it.tag_id = t.id
                JOIN slides sl ON i.slide_id = sl.id
                JOIN topics tp ON sl.topic_id = tp.id
                WHERE t.name = @tagName
                ORDER BY i.name
                LIMIT @limit OFFSET @offset";

            var results = new List<IllustrationItem>();
            using (var cmd = new SQLiteCommand(sql, _conn))
            {
                cmd.Parameters.AddWithValue("@tagName", tagName);
                cmd.Parameters.AddWithValue("@limit", limit);
                cmd.Parameters.AddWithValue("@offset", offset);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        results.Add(ReadIllustration(reader));
                    }
                }
            }
            return results;
        }

        public TilingMetadata GetTilingMetadata(int illustrationId)
        {
            const string sql = @"
                SELECT it.metadata
                FROM illustration_tags it
                JOIN tags t ON it.tag_id = t.id
                WHERE it.illustration_id = @id
                  AND t.name IN ('tile-horizontal', 'tile-vertical', 'tileable-2d')
                  AND it.metadata IS NOT NULL AND it.metadata != '{}'
                LIMIT 1";

            using (var cmd = new SQLiteCommand(sql, _conn))
            {
                cmd.Parameters.AddWithValue("@id", illustrationId);
                var result = cmd.ExecuteScalar() as string;
                if (string.IsNullOrEmpty(result))
                {
                    return null;
                }
                return JsonConvert.DeserializeObject<TilingMetadata>(result);
            }
        }

        public List<TopicInfo> GetTopics()
        {
            const string sql = "SELECT id, name, illustration_count FROM topics ORDER BY name";
            var results = new List<TopicInfo>();
            using (var cmd = new SQLiteCommand(sql, _conn))
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    results.Add(new TopicInfo
                    {
                        Id = reader.GetInt32(0),
                        Name = reader.GetString(1),
                        IllustrationCount = reader.GetInt32(2)
                    });
                }
            }
            return results;
        }

        public List<TagInfo> GetTags()
        {
            const string sql = @"
                SELECT t.id, t.name, COUNT(*) as cnt
                FROM illustration_tags it
                JOIN tags t ON it.tag_id = t.id
                GROUP BY t.name
                ORDER BY cnt DESC";

            var results = new List<TagInfo>();
            using (var cmd = new SQLiteCommand(sql, _conn))
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    results.Add(new TagInfo
                    {
                        Id = reader.GetInt32(0),
                        Name = reader.GetString(1),
                        Count = reader.GetInt32(2)
                    });
                }
            }
            return results;
        }

        public string ResolveAssetPath(string relativePath)
        {
            string fullPath = Path.Combine(_basePath, relativePath);
            if (File.Exists(fullPath))
            {
                return fullPath;
            }
            string withoutSuffix = System.Text.RegularExpressions.Regex.Replace(
                Path.GetFileNameWithoutExtension(fullPath), @"_\d+$", "");
            string fallback = Path.Combine(
                Path.GetDirectoryName(fullPath),
                withoutSuffix + Path.GetExtension(fullPath));
            return File.Exists(fallback) ? fallback : fullPath;
        }

        public void Dispose()
        {
            if (_conn != null)
            {
                _conn.Close();
                _conn.Dispose();
            }
        }

        private bool CheckFts5Support()
        {
            try
            {
                using (var cmd = new SQLiteCommand("SELECT count(*) FROM search LIMIT 1", _conn))
                {
                    cmd.ExecuteScalar();
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private List<IllustrationItem> SearchFts5(string query, int limit)
        {
            const string sql = @"
                SELECT i.id, i.name, i.description, i.pptx_file, i.pptx_slide, i.pptx_shape_index,
                       i.svg_path, i.png_path, i.width, i.height,
                       tp.name as topic, sl.title as slide_title
                FROM search s
                JOIN illustrations i ON i.id = s.rowid
                JOIN slides sl ON i.slide_id = sl.id
                JOIN topics tp ON sl.topic_id = tp.id
                WHERE search MATCH @query
                ORDER BY rank
                LIMIT @limit";

            var results = new List<IllustrationItem>();
            using (var cmd = new SQLiteCommand(sql, _conn))
            {
                cmd.Parameters.AddWithValue("@query", query);
                cmd.Parameters.AddWithValue("@limit", limit);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        results.Add(ReadIllustration(reader));
                    }
                }
            }

            return results;
        }

        private List<IllustrationItem> SearchLike(string query, int limit)
        {
            string pattern = "%" + query.Replace(" ", "%") + "%";
            const string sql = @"
                SELECT i.id, i.name, i.description, i.pptx_file, i.pptx_slide, i.pptx_shape_index,
                       i.svg_path, i.png_path, i.width, i.height,
                       tp.name as topic, sl.title as slide_title
                FROM illustrations i
                JOIN slides sl ON i.slide_id = sl.id
                JOIN topics tp ON sl.topic_id = tp.id
                WHERE i.name LIKE @pattern
                   OR i.description LIKE @pattern
                   OR tp.name LIKE @pattern
                ORDER BY i.name
                LIMIT @limit";

            var results = new List<IllustrationItem>();
            using (var cmd = new SQLiteCommand(sql, _conn))
            {
                cmd.Parameters.AddWithValue("@pattern", pattern);
                cmd.Parameters.AddWithValue("@limit", limit);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        results.Add(ReadIllustration(reader));
                    }
                }
            }

            return results;
        }

        private IllustrationItem ReadIllustration(SQLiteDataReader reader)
        {
            return new IllustrationItem
            {
                Id = reader.GetInt32(reader.GetOrdinal("id")),
                Name = reader.GetString(reader.GetOrdinal("name")),
                Description = reader.IsDBNull(reader.GetOrdinal("description")) ? "" : reader.GetString(reader.GetOrdinal("description")),
                PptxFile = reader.IsDBNull(reader.GetOrdinal("pptx_file")) ? "" : reader.GetString(reader.GetOrdinal("pptx_file")),
                PptxSlide = reader.IsDBNull(reader.GetOrdinal("pptx_slide")) ? 0 : reader.GetInt32(reader.GetOrdinal("pptx_slide")),
                PptxShapeIndex = reader.IsDBNull(reader.GetOrdinal("pptx_shape_index")) ? 0 : reader.GetInt32(reader.GetOrdinal("pptx_shape_index")),
                SvgPath = reader.IsDBNull(reader.GetOrdinal("svg_path")) ? "" : reader.GetString(reader.GetOrdinal("svg_path")),
                PngPath = reader.IsDBNull(reader.GetOrdinal("png_path")) ? "" : reader.GetString(reader.GetOrdinal("png_path")),
                Width = reader.IsDBNull(reader.GetOrdinal("width")) ? 0 : (float)reader.GetDouble(reader.GetOrdinal("width")),
                Height = reader.IsDBNull(reader.GetOrdinal("height")) ? 0 : (float)reader.GetDouble(reader.GetOrdinal("height")),
                Topic = reader.GetString(reader.GetOrdinal("topic")),
                SlideTitle = reader.IsDBNull(reader.GetOrdinal("slide_title")) ? "" : reader.GetString(reader.GetOrdinal("slide_title"))
            };
        }
    }
}
