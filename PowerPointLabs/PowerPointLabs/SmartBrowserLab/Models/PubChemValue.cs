using Newtonsoft.Json;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class PubChemValue
    {
        [JsonProperty("ival")]
        public int? IntValue { get; set; }

        [JsonProperty("fval")]
        public double? FloatValue { get; set; }

        [JsonProperty("sval")]
        public string StringValue { get; set; }

        [JsonProperty("fvec")]
        public double[] FloatVector { get; set; }
    }
}
