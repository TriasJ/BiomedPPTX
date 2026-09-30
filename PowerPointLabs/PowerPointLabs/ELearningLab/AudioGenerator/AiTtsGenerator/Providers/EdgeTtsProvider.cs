using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;

using Newtonsoft.Json;

namespace PowerPointLabs.ELearningLab.AudioGenerator.AiTtsGenerator.Providers
{
    public class EdgeTtsProvider : IAiTtsProvider
    {
        private const string VoiceListUrl =
            "https://speech.platform.bing.com/consumer/speech/synthesize/readaloud/voices/list" +
            "?trustedclienttoken=6A5AA1D4EAFF4E9FB37E23D68491D6F4";

        private List<AiTtsVoice> _cachedVoices;

        public string Name
        {
            get { return "Edge TTS (Free)"; }
        }

        public bool RequiresApiKey
        {
            get { return false; }
        }

        public bool IsConfigured
        {
            get { return true; }
        }

        public void Configure(string apiKey, string endpoint)
        {
        }

        public List<AiTtsVoice> GetVoices()
        {
            if (_cachedVoices != null)
            {
                return _cachedVoices;
            }

            _cachedVoices = new List<AiTtsVoice>();
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                    | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                using (WebClient client = new WebClient())
                {
                    client.Headers.Add("User-Agent", "Mozilla/5.0");
                    string json = client.DownloadString(VoiceListUrl);
                    List<EdgeVoiceInfo> voices = JsonConvert.DeserializeObject<List<EdgeVoiceInfo>>(json);
                    if (voices != null)
                    {
                        foreach (EdgeVoiceInfo v in voices)
                        {
                            _cachedVoices.Add(new AiTtsVoice(
                                v.FriendlyName,
                                v.ShortName,
                                v.Locale,
                                v.Gender,
                                "Edge TTS"));
                        }
                    }
                }
            }
            catch (Exception)
            {
                _cachedVoices.Add(new AiTtsVoice(
                    "Aria (US English)", "en-US-AriaNeural", "en-US", "Female", "Edge TTS"));
                _cachedVoices.Add(new AiTtsVoice(
                    "Guy (US English)", "en-US-GuyNeural", "en-US", "Male", "Edge TTS"));
                _cachedVoices.Add(new AiTtsVoice(
                    "Dalia (Mexican Spanish)", "es-MX-DaliaNeural", "es-MX", "Female", "Edge TTS"));
                _cachedVoices.Add(new AiTtsVoice(
                    "Xiaoxiao (Chinese)", "zh-CN-XiaoxiaoNeural", "zh-CN", "Female", "Edge TTS"));
            }

            return _cachedVoices;
        }

        public void Synthesize(string text, string voiceName, string outputFilePath)
        {
            string tempMp3 = Path.Combine(Path.GetTempPath(),
                "biomedpptx_tts_" + DateTime.Now.Ticks.ToString() + ".mp3");

            try
            {
                string escapedText = text.Replace("\"", "'").Replace("\r", " ").Replace("\n", " ");

                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = "python",
                    Arguments = string.Format(
                        "-m edge_tts --text \"{0}\" --voice {1} --write-media \"{2}\"",
                        escapedText,
                        voiceName,
                        tempMp3),
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                Process process = Process.Start(startInfo);
                if (process != null)
                {
                    process.WaitForExit(30000);
                }

                if (File.Exists(tempMp3) && new FileInfo(tempMp3).Length > 0)
                {
                    ConvertMp3ToWav(tempMp3, outputFilePath);
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                if (File.Exists(tempMp3))
                {
                    try
                    {
                        File.Delete(tempMp3);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static void ConvertMp3ToWav(string mp3Path, string wavPath)
        {
            using (NAudio.Wave.Mp3FileReader reader = new NAudio.Wave.Mp3FileReader(mp3Path))
            {
                NAudio.Wave.WaveFileWriter.CreateWaveFile(wavPath, reader);
            }
        }
    }

}
