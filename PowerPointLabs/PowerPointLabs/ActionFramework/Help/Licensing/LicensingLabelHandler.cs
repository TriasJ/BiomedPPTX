using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Help
{
    [ExportLabelRibbonId(HelpText.LicensingTag)]
    class LicensingLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return HelpText.LicensingButtonLabel;
        }
    }
}
