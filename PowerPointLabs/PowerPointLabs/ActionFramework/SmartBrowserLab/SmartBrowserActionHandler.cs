using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.SmartBrowserLab
{
    [ExportActionRibbonId(SmartBrowserLabText.PaneTag)]
    class SmartBrowserActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            // TODO: Register and toggle the SMART Browser task pane
            // This will be implemented when the WPF pane is created
            System.Windows.Forms.MessageBox.Show(
                "SMART Browser is under construction.\n\n" +
                "This will open a searchable browser for 4,474 biomedical illustrations " +
                "from the SMART Medical Illustrations library and NCBI BioArt.",
                "BiomedPPTX - SMART Browser",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }
    }
}
