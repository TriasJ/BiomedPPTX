using System.Collections.Generic;

namespace PowerPointLabs.ELearningLab.AudioGenerator.AiTtsGenerator
{
    public interface IAiTtsProvider
    {
        string Name { get; }

        bool RequiresApiKey { get; }

        bool IsConfigured { get; }

        void Configure(string apiKey, string endpoint);

        List<AiTtsVoice> GetVoices();

        void Synthesize(string text, string voiceName, string outputFilePath);
    }
}
