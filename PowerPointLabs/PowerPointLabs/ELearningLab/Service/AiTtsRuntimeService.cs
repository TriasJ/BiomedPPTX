using System.Collections.Generic;

using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.AudioGenerator.AiTtsGenerator;

namespace PowerPointLabs.ELearningLab.Service
{
    public static class AiTtsRuntimeService
    {
        private static List<IVoice> _voices;

        public static IReadOnlyList<IVoice> Voices
        {
            get
            {
                if (_voices == null)
                {
                    _voices = new List<IVoice>();
                    IAiTtsProvider provider = AiTtsProviderRegistry.GetActiveProvider();
                    if (provider != null)
                    {
                        foreach (AiTtsVoice v in provider.GetVoices())
                        {
                            _voices.Add(v);
                        }
                    }
                }

                return _voices.AsReadOnly();
            }
        }

        public static void SaveStringToWaveFile(string text, string filePath, AiTtsVoice voice)
        {
            IAiTtsProvider provider = AiTtsProviderRegistry.GetActiveProvider();
            if (provider != null && voice != null)
            {
                provider.Synthesize(text, voice.Id, filePath);
            }
        }

        public static void RefreshVoices()
        {
            _voices = null;
        }
    }
}
