using Newtonsoft.Json;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class PubChemCompoundData
    {
        [JsonProperty("id")]
        public PubChemCompoundId Id { get; set; }

        [JsonProperty("atoms")]
        public PubChemAtoms Atoms { get; set; }

        [JsonProperty("bonds")]
        public PubChemBonds Bonds { get; set; }

        [JsonProperty("coords")]
        public PubChemCoords[] Coords { get; set; }

        [JsonProperty("charge")]
        public int Charge { get; set; }

        [JsonProperty("props")]
        public PubChemProperty[] Props { get; set; }

        [JsonProperty("count")]
        public PubChemCount Count { get; set; }
    }
}
