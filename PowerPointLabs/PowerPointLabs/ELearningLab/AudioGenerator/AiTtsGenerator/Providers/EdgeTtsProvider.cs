using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.WebSockets;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using Newtonsoft.Json;

namespace PowerPointLabs.ELearningLab.AudioGenerator.AiTtsGenerator.Providers
{
    public class EdgeTtsProvider : IAiTtsProvider
    {
        private const string TrustedToken = "6A5AA1D4EAFF4E9FB37E23D68491D6F4";

        private const string VoiceListUrl =
            "https://speech.platform.bing.com/consumer/speech/synthesize/readaloud/voices/list"
            + "?trustedclienttoken=" + TrustedToken;

        private const string WsEndpoint =
            "wss://speech.platform.bing.com/consumer/speech/synthesize/readaloud/edge/v1";

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
                            string name = v.FriendlyName != null ? v.FriendlyName : v.ShortName;
                            _cachedVoices.Add(new AiTtsVoice(
                                name, v.ShortName, v.Locale, v.Gender, "Edge TTS"));
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
                    "Jenny (US English)", "en-US-JennyNeural", "en-US", "Female", "Edge TTS"));
                _cachedVoices.Add(new AiTtsVoice(
                    "Dalia (Mexican Spanish)", "es-MX-DaliaNeural", "es-MX", "Female", "Edge TTS"));
                _cachedVoices.Add(new AiTtsVoice(
                    "Jorge (Mexican Spanish)", "es-MX-JorgeNeural", "es-MX", "Male", "Edge TTS"));
            }

            return _cachedVoices;
        }

        public void Synthesize(string text, string voiceName, string outputFilePath)
        {
            Log("=== Synthesize START ===");
            Log("Text: " + (text != null ? text.Substring(0, Math.Min(text.Length, 80)) : "NULL"));
            Log("Voice: " + voiceName);
            Log("Output: " + outputFilePath);

            string tempMp3 = Path.Combine(Path.GetTempPath(),
                "biomedpptx_tts_" + Guid.NewGuid().ToString("N") + ".mp3");

            try
            {
                Log("Trying WebSocket...");
                Task<byte[]> task = Task.Run(() => SynthesizeViaWebSocket(text, voiceName));
                byte[] audioData = task.GetAwaiter().GetResult();

                if (audioData != null && audioData.Length > 0)
                {
                    File.WriteAllBytes(tempMp3, audioData);
                    Log("WebSocket OK: " + audioData.Length + " bytes");
                    ConvertMp3ToWav(tempMp3, outputFilePath);
                    Log("WAV written: " + (File.Exists(outputFilePath) ? new FileInfo(outputFilePath).Length + " bytes" : "MISSING"));
                    return;
                }

                Log("WebSocket returned no data");
            }
            catch (Exception ex)
            {
                Log("WebSocket failed: " + ex.GetType().Name + ": " + ex.Message);
            }

            try
            {
                Log("Falling back to edge-tts CLI...");
                SynthesizeViaCli(text, voiceName, tempMp3);

                if (File.Exists(tempMp3) && new FileInfo(tempMp3).Length > 0)
                {
                    Log("CLI MP3: " + new FileInfo(tempMp3).Length + " bytes");
                    ConvertMp3ToWav(tempMp3, outputFilePath);
                    Log("WAV written: " + (File.Exists(outputFilePath) ? new FileInfo(outputFilePath).Length + " bytes" : "MISSING"));
                    return;
                }

                Log("CLI produced no audio");
            }
            catch (Exception ex2)
            {
                Log("CLI also failed: " + ex2.GetType().Name + ": " + ex2.Message);
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

            Log("=== Synthesize END ===");
        }

        private static void SynthesizeViaCli(string text, string voiceName, string outputMp3)
        {
            string escapedText = text.Replace("\"", "'").Replace("\r", " ").Replace("\n", " ");
            string args = string.Format(
                "-m edge_tts --text \"{0}\" --voice {1} --write-media \"{2}\"",
                escapedText, voiceName, outputMp3);

            Log("CLI args: python " + args.Substring(0, System.Math.Min(args.Length, 100)));

            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "python",
                Arguments = args,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using (System.Diagnostics.Process proc = System.Diagnostics.Process.Start(startInfo))
            {
                string stderr = proc.StandardError.ReadToEnd();
                proc.WaitForExit(30000);

                Log("CLI exit code: " + proc.ExitCode);
                if (!string.IsNullOrEmpty(stderr))
                {
                    Log("CLI stderr: " + stderr.Substring(0, System.Math.Min(stderr.Length, 200)));
                }
            }
        }

        private static string _logPath = Path.Combine(Path.GetTempPath(), "BiomedPPTX_EdgeTTS_Debug.txt");

        private static void Log(string message)
        {
            try
            {
                File.AppendAllText(_logPath,
                    DateTime.Now.ToString("HH:mm:ss.fff") + " " + message + "\r\n");
            }
            catch (Exception)
            {
            }
        }

        private static string GenerateSecMsGec()
        {
            long winEpoch = 11644473600L;
            double unixTime = (DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds;
            double ticks = unixTime + winEpoch;
            ticks -= ticks % 300;
            ticks *= 1e9 / 100;

            string input = string.Format("{0:F0}{1}", ticks, TrustedToken);

            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.ASCII.GetBytes(input));
                StringBuilder sb = new StringBuilder();
                foreach (byte b in hash)
                {
                    sb.Append(b.ToString("X2"));
                }

                return sb.ToString();
            }
        }

        private static string GenerateMuid()
        {
            byte[] bytes = new byte[16];
            using (var rng = new System.Security.Cryptography.RNGCryptoServiceProvider())
            {
                rng.GetBytes(bytes);
            }

            StringBuilder sb = new StringBuilder();
            foreach (byte b in bytes)
            {
                sb.Append(b.ToString("X2"));
            }

            return sb.ToString();
        }

        private static string BuildWsUrl(string connectionId)
        {
            string secGec = GenerateSecMsGec();
            return WsEndpoint
                + "?TrustedClientToken=" + TrustedToken
                + "&Sec-MS-GEC=" + secGec
                + "&Sec-MS-GEC-Version=1-143.0.3650.75"
                + "&ConnectionId=" + connectionId;
        }

        private static async Task<byte[]> SynthesizeViaWebSocket(string text, string voiceName)
        {
            string connectionId = Guid.NewGuid().ToString("N");
            string wsUrl = BuildWsUrl(connectionId);

            Log("WebSocket connecting to: " + wsUrl.Substring(0, 80) + "...");

            using (ClientWebSocket ws = new ClientWebSocket())
            {
                try
                {
                    ws.Options.SetRequestHeader("Pragma", "no-cache");
                    ws.Options.SetRequestHeader("Cache-Control", "no-cache");
                    ws.Options.SetRequestHeader("Origin",
                        "chrome-extension://jdiccldimpdaibmpdkjnbmckianbfold");
                    ws.Options.SetRequestHeader("Accept-Encoding", "gzip, deflate, br");
                    ws.Options.SetRequestHeader("Accept-Language", "en-US,en;q=0.9");
                    ws.Options.SetRequestHeader("Cookie",
                        "muid=" + GenerateMuid() + ";");
                }
                catch (Exception)
                {
                }

                CancellationTokenSource cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                await ws.ConnectAsync(new Uri(wsUrl), cts.Token).ConfigureAwait(false);
                Log("WebSocket connected, state: " + ws.State);

                string timestamp = DateTime.UtcNow.ToString("ddd MMM dd yyyy HH:mm:ss");

                string configMsg = "X-Timestamp:" + timestamp + "\r\n"
                    + "Content-Type:application/json; charset=utf-8\r\n"
                    + "Path:speech.config\r\n\r\n"
                    + "{\"context\":{\"synthesis\":{\"audio\":{\"metadataoptions\":"
                    + "{\"sentenceBoundaryEnabled\":\"false\",\"wordBoundaryEnabled\":\"false\"},"
                    + "\"outputFormat\":\"audio-24khz-48kbitrate-mono-mp3\"}}}}";

                byte[] configBytes = Encoding.UTF8.GetBytes(configMsg);
                await ws.SendAsync(
                    new ArraySegment<byte>(configBytes),
                    WebSocketMessageType.Text, true, cts.Token).ConfigureAwait(false);

                string escapedText = text
                    .Replace("&", "&amp;")
                    .Replace("<", "&lt;")
                    .Replace(">", "&gt;")
                    .Replace("\r", " ")
                    .Replace("\n", " ");

                string ssml = "<speak version=\"1.0\" xmlns=\"http://www.w3.org/2001/10/synthesis\" xml:lang=\"en-US\">"
                    + "<voice name=\"" + voiceName + "\">"
                    + "<prosody pitch=\"+0Hz\" rate=\"+0%\" volume=\"+0%\">"
                    + escapedText
                    + "</prosody></voice></speak>";

                string ssmlMsg = "X-RequestId:" + connectionId + "\r\n"
                    + "Content-Type:application/ssml+xml\r\n"
                    + "X-Timestamp:" + timestamp + "\r\n"
                    + "Path:ssml\r\n\r\n" + ssml;

                byte[] ssmlBytes = Encoding.UTF8.GetBytes(ssmlMsg);
                await ws.SendAsync(
                    new ArraySegment<byte>(ssmlBytes),
                    WebSocketMessageType.Text, true, cts.Token).ConfigureAwait(false);

                byte[] separatorBytes = Encoding.UTF8.GetBytes("\r\n\r\n");

                using (MemoryStream audioStream = new MemoryStream())
                {
                    byte[] buffer = new byte[8192];
                    bool done = false;

                    while (!done && ws.State == WebSocketState.Open)
                    {
                        WebSocketReceiveResult result = await ws.ReceiveAsync(
                            new ArraySegment<byte>(buffer), cts.Token).ConfigureAwait(false);

                        if (result.MessageType == WebSocketMessageType.Binary)
                        {
                            int sepIndex = FindBytes(buffer, separatorBytes, 0, result.Count);
                            if (sepIndex >= 0)
                            {
                                int audioStart = sepIndex + separatorBytes.Length;
                                int audioLen = result.Count - audioStart;
                                if (audioLen > 0)
                                {
                                    audioStream.Write(buffer, audioStart, audioLen);
                                }
                            }
                            else
                            {
                                audioStream.Write(buffer, 0, result.Count);
                            }

                            while (!result.EndOfMessage)
                            {
                                result = await ws.ReceiveAsync(
                                    new ArraySegment<byte>(buffer), cts.Token).ConfigureAwait(false);
                                audioStream.Write(buffer, 0, result.Count);
                            }
                        }
                        else if (result.MessageType == WebSocketMessageType.Text)
                        {
                            StringBuilder textBuilder = new StringBuilder();
                            textBuilder.Append(Encoding.UTF8.GetString(buffer, 0, result.Count));
                            while (!result.EndOfMessage)
                            {
                                result = await ws.ReceiveAsync(
                                    new ArraySegment<byte>(buffer), cts.Token).ConfigureAwait(false);
                                textBuilder.Append(Encoding.UTF8.GetString(buffer, 0, result.Count));
                            }

                            if (textBuilder.ToString().Contains("Path:turn.end"))
                            {
                                done = true;
                            }
                        }
                        else if (result.MessageType == WebSocketMessageType.Close)
                        {
                            done = true;
                        }
                    }

                    return audioStream.ToArray();
                }
            }
        }

        private static int FindBytes(byte[] haystack, byte[] needle, int start, int length)
        {
            int end = start + length - needle.Length;
            for (int i = start; i <= end; i++)
            {
                bool found = true;
                for (int j = 0; j < needle.Length; j++)
                {
                    if (haystack[i + j] != needle[j])
                    {
                        found = false;
                        break;
                    }
                }

                if (found)
                {
                    return i;
                }
            }

            return -1;
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
