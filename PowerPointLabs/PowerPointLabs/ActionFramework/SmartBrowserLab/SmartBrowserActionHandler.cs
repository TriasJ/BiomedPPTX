using System;
using System.IO;

using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.SmartBrowserLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.SmartBrowserLab
{
    [ExportActionRibbonId(SmartBrowserLabText.PaneTag)]
    class SmartBrowserActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.RegisterTaskPane(typeof(SmartBrowserPane), SmartBrowserLabText.TaskPanelTitle);
            CustomTaskPane browserPane = this.GetTaskPane(typeof(SmartBrowserPane));

            if (browserPane == null)
            {
                return;
            }

            browserPane.Visible = !browserPane.Visible;

            if (browserPane.Visible)
            {
                SmartBrowserPane pane = browserPane.Control as SmartBrowserPane;
                if (pane != null)
                {
                    string assetsPath = Path.Combine(ThisAddIn.AppDataFolder, "Assets", "SMART-Library");
                    string dbPath = Path.Combine(assetsPath, "illustrations.db");

                    if (!File.Exists(dbPath))
                    {
                        assetsPath = FindSmartLibraryPath();
                        dbPath = Path.Combine(assetsPath, "illustrations.db");
                    }

                    if (File.Exists(dbPath))
                    {
                        pane.InitBrowser(dbPath, assetsPath);
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show(
                            "SMART-Library database not found.\n\n" +
                            "Expected location: " + dbPath + "\n\n" +
                            "Please ensure the SMART-Library assets are installed.",
                            "BiomedPPTX",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                    }
                }
            }
        }

        private string FindSmartLibraryPath()
        {
            string[] searchPaths = new[]
            {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "SMART-Library"),
                Path.Combine(ThisAddIn.AppDataFolder, "Assets", "SMART-Library"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "__Scratch", "SMART-Library")
            };

            foreach (string path in searchPaths)
            {
                if (File.Exists(Path.Combine(path, "illustrations.db")))
                {
                    return path;
                }
            }

            return searchPaths[0];
        }
    }
}
