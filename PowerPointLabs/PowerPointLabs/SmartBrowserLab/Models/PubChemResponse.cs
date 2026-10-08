using Newtonsoft.Json;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class PubChemResponse
    {
        [JsonProperty("PC_Compounds")]
        public PubChemCompoundData[] Compounds { get; set; }
    }
}
