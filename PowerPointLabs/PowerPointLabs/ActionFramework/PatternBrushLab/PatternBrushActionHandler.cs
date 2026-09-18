using System;
using System.IO;

using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.PatternBrushLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PatternBrushLab
{
    [ExportActionRibbonId(PatternBrushLabText.PaneTag)]
    class PatternBrushActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.RegisterTaskPane(typeof(PatternBrushPane), PatternBrushLabText.TaskPanelTitle);
            CustomTaskPane brushPane = this.GetTaskPane(typeof(PatternBrushPane));

            if (brushPane == null)
            {
                return;
            }

            brushPane.Visible = !brushPane.Visible;

            if (brushPane.Visible)
            {
                PatternBrushPane pane = brushPane.Control as PatternBrushPane;
                if (pane != null)
                {
                    string assetsPath = Path.Combine(ThisAddIn.AppDataFolder, "Assets", "SMART-Library");
                    string dbPath = Path.Combine(assetsPath, "illustrations.db");

                    if (!File.Exists(dbPath))
                    {
                        string[] searchPaths = new[]
                        {
                            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "__Scratch", "SMART-Library"),
                            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "SMART-Library")
                        };

                        foreach (string path in searchPaths)
                        {
                            string candidate = Path.Combine(path, "illustrations.db");
                            if (File.Exists(candidate))
                            {
                                assetsPath = path;
                                dbPath = candidate;
                                break;
                            }
                        }
                    }

                    if (File.Exists(dbPath))
                    {
                        pane.InitBrush(dbPath, assetsPath);
                    }
                }
            }
        }
    }
}
