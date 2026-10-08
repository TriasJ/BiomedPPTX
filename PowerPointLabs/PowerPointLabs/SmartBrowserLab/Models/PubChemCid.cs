using Newtonsoft.Json;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class PubChemCid
    {
        [JsonProperty("cid")]
        public int Cid { get; set; }
    }
}
