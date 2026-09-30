using Newtonsoft.Json;

namespace PowerPointLabs.ELearningLab.AudioGenerator.AiTtsGenerator.Providers
{
    internal class EdgeVoiceInfo
    {
        [JsonProperty("ShortName")]
        public string ShortName { get; set; }

        [JsonProperty("FriendlyName")]
        public string FriendlyName { get; set; }

        [JsonProperty("Locale")]
        public string Locale { get; set; }

        [JsonProperty("Gender")]
        public string Gender { get; set; }
    }
}
