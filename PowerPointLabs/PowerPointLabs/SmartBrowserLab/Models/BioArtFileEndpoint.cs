using Newtonsoft.Json;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class BioArtFileEndpoint
    {
        [JsonProperty("file_id")]
        public string FileId { get; set; }

        [JsonProperty("url")]
        public string FileUrl { get; set; }
    }
}
