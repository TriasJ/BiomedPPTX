using System.Collections.Generic;

using PowerPointLabs.ELearningLab.AudioGenerator.AiTtsGenerator.Providers;

namespace PowerPointLabs.ELearningLab.AudioGenerator.AiTtsGenerator
{
    public static class AiTtsProviderRegistry
    {
        private static List<IAiTtsProvider> _providers = new List<IAiTtsProvider>();
        private static IAiTtsProvider _activeProvider;

        static AiTtsProviderRegistry()
        {
            RegisterProvider(new EdgeTtsProvider());
        }

        public static void RegisterProvider(IAiTtsProvider provider)
        {
            _providers.Add(provider);
            if (_activeProvider == null && provider.IsConfigured)
            {
                _activeProvider = provider;
            }
        }

        public static List<IAiTtsProvider> GetProviders()
        {
            return new List<IAiTtsProvider>(_providers);
        }

        public static IAiTtsProvider GetActiveProvider()
        {
            return _activeProvider;
        }

        public static void SetActiveProvider(IAiTtsProvider provider)
        {
            _activeProvider = provider;
        }

        public static IAiTtsProvider GetProviderByName(string name)
        {
            foreach (IAiTtsProvider p in _providers)
            {
                if (p.Name == name)
                {
                    return p;
                }
            }

            return null;
        }
    }
}
