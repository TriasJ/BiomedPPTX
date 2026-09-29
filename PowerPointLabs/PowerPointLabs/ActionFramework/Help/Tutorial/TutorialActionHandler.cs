using System.Diagnostics;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Help
{
    [ExportActionRibbonId(HelpText.TutorialTag)]
    class TutorialActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            string[] searchPaths = new string[]
            {
                System.IO.Path.Combine(ThisAddIn.AppDataFolder, CommonText.QuickTutorialFileName),
                System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, CommonText.QuickTutorialFileName),
                System.IO.Path.Combine(System.IO.Path.GetDirectoryName(
                    System.Reflection.Assembly.GetExecutingAssembly().Location), CommonText.QuickTutorialFileName)
            };

            string tutorialPath = null;
            foreach (string path in searchPaths)
            {
                if (System.IO.File.Exists(path))
                {
                    tutorialPath = path;
                    break;
                }
            }

            try
            {
                if (tutorialPath != null)
                {
                    Process.Start("POWERPNT", "\"" + tutorialPath + "\"");
                }
                else
                {
                    Process.Start(CommonText.HelpDocumentUrl);
                }
            }
            catch (System.Exception)
            {
                Logger.Log("TutorialButtonClick: Failed to open tutorial file!", ActionFramework.Common.Logger.LogType.Error);
            }
        }
    }
}
