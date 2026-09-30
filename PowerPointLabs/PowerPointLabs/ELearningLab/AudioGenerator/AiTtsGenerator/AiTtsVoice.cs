namespace PowerPointLabs.ELearningLab.AudioGenerator.AiTtsGenerator
{
    public class AiTtsVoice : IVoice
    {
        private string _voiceName;

        public AiTtsVoice(string name, string id, string language, string gender, string provider)
        {
            _voiceName = name;
            Id = id;
            Language = language;
            Gender = gender;
            Provider = provider;
        }

        public string Id { get; set; }

        public string Language { get; set; }

        public string Gender { get; set; }

        public string Provider { get; set; }

        public string Voice
        {
            get
            {
                return _voiceName;
            }
        }

        public override string VoiceName
        {
            get
            {
                return _voiceName;
            }
        }

        public override object Clone()
        {
            return new AiTtsVoice(_voiceName, Id, Language, Gender, Provider);
        }
    }
}
