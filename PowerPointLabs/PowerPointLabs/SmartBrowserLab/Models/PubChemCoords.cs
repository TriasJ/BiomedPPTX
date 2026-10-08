using Newtonsoft.Json;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class PubChemCoords
    {
        [JsonProperty("type")]
        public int[] Type { get; set; }

        [JsonProperty("aid")]
        public int[] AtomIds { get; set; }

        [JsonProperty("conformers")]
        public PubChemConformer[] Conformers { get; set; }
    }
}
