using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Views;

namespace PowerPointLabs.ActionFramework.Help
{
    [ExportActionRibbonId(HelpText.LicensingTag)]
    class LicensingActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            LicensingDialogBox dialog = new LicensingDialogBox();
            dialog.ShowDialog();
        }
    }
}
