using Newtonsoft.Json;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class PubChemCompoundId
    {
        [JsonProperty("id")]
        public PubChemCid CidWrapper { get; set; }

        public int GetCid()
        {
            if (CidWrapper != null)
            {
                return CidWrapper.Cid;
            }

            return 0;
        }
    }
}
