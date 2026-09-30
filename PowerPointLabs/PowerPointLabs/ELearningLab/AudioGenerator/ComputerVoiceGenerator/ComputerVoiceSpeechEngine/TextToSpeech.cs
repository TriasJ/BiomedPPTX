using System.Collections.Generic;
using System.Linq;
using System.Speech.Synthesis;

using PowerPointLabs.ELearningLab.AudioGenerator.AiTtsGenerator;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.Models;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    static class TextToSpeech
    {
        public static IEnumerable<string> GetVoices()
        {
            using (SpeechSynthesizer synthesizer = new SpeechSynthesizer())
            {
                System.Collections.ObjectModel.ReadOnlyCollection<InstalledVoice> installedVoices = synthesizer.GetInstalledVoices();
                IEnumerable<InstalledVoice> voices = installedVoices.Where(voice => voice.Enabled);
                return voices.Select(voice => voice.VoiceInfo.Name);
            }
        }

        public static void SaveStringToWaveFiles(string notesText, string folderPath, string fileNameFormat)
        {
            TaggedText taggedNotes = new TaggedText(notesText);
            List<string> stringsToSave = taggedNotes.SplitByClicks();
            //MD5 md5 = MD5.Create();

            for (int i = 0; i < stringsToSave.Count; i++)
            {
                string textToSave = stringsToSave[i];
                string baseFileName = string.Format(fileNameFormat, i + 1);

                // The first item will autoplay; everything else is triggered by a click.
                string fileName = i > 0 ? baseFileName + " (OnClick)" : baseFileName;
                string filePath = folderPath + "\\" + fileName + ".wav";

                switch (AudioSettingService.selectedVoiceType)
                {
                    case VoiceType.ComputerVoice:
                        ComputerVoiceRuntimeService.SaveStringToWaveFile(textToSave, filePath,
                            AudioSettingService.selectedVoice as ComputerVoice);
                        break;
                    case VoiceType.AzureVoice:
                        AzureRuntimeService.SaveStringToWaveFileWithAzureVoice(textToSave, filePath, 
                            AudioSettingService.selectedVoice as AzureVoice);
                        break;
                    case VoiceType.WatsonVoice:
                        WatsonRuntimeService.SaveStringToWaveFile(textToSave, filePath,
                            AudioSettingService.selectedVoice as WatsonVoice);
                        break;
                    case VoiceType.AiTtsVoice:
                        try
                        {
                            System.IO.File.AppendAllText(
                                System.IO.Path.Combine(System.IO.Path.GetTempPath(), "BiomedPPTX_EdgeTTS_Debug.txt"),
                                System.DateTime.Now.ToString("HH:mm:ss") + " TextToSpeech routing to AiTts: voice="
                                + (AudioSettingService.selectedVoice != null ? AudioSettingService.selectedVoice.VoiceName : "NULL")
                                + " text=" + textToSave.Substring(0, System.Math.Min(textToSave.Length, 50))
                                + " file=" + filePath + "\r\n");
                        }
                        catch (System.Exception)
                        {
                        }

                        AiTtsRuntimeService.SaveStringToWaveFile(textToSave, filePath,
                            AudioSettingService.selectedVoice as AiTtsVoice);
                        break;
                    default:
                        break;
                }
            }
        }
    }
}