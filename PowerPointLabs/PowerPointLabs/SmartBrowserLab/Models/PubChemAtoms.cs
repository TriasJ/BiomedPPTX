using Newtonsoft.Json;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class PubChemAtoms
    {
        [JsonProperty("aid")]
        public int[] AtomIds { get; set; }

        [JsonProperty("element")]
        public int[] Elements { get; set; }
    }
}
