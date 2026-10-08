using Newtonsoft.Json;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class PubChemConformer
    {
        [JsonProperty("x")]
        public double[] X { get; set; }

        [JsonProperty("y")]
        public double[] Y { get; set; }

        [JsonProperty("z")]
        public double[] Z { get; set; }

        [JsonProperty("data")]
        public PubChemConformerData[] Data { get; set; }
    }
}
