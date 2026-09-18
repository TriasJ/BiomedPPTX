using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PatternBrushLab
{
    [ExportActionRibbonId(PatternBrushLabText.PaneTag)]
    class PatternBrushActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            // TODO: Register and toggle the Pattern Brush task pane
            System.Windows.Forms.MessageBox.Show(
                "Pattern Brush is under construction.\n\n" +
                "This tool will transform lines and freeforms into tiled biological patterns " +
                "(phospholipid bilayers, epithelial tissue, muscle fibers, vascular stents).",
                "BiomedPPTX - Pattern Brush",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }
    }
}
