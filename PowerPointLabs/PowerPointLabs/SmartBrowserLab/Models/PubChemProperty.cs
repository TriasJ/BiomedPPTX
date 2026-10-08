using Newtonsoft.Json;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class PubChemProperty
    {
        [JsonProperty("urn")]
        public PubChemUrn Urn { get; set; }

        [JsonProperty("value")]
        public PubChemValue Value { get; set; }
    }
}
