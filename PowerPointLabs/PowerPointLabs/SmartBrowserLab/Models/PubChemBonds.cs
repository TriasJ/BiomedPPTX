using Newtonsoft.Json;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class PubChemBonds
    {
        [JsonProperty("aid1")]
        public int[] AtomId1 { get; set; }

        [JsonProperty("aid2")]
        public int[] AtomId2 { get; set; }

        [JsonProperty("order")]
        public int[] Order { get; set; }
    }
}
