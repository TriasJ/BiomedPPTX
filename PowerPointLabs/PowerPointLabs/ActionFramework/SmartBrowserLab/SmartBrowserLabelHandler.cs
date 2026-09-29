using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.SmartBrowserLab
{
    [ExportLabelRibbonId(SmartBrowserLabText.PaneTag)]
    class SmartBrowserLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return SmartBrowserLabText.ButtonLabel;
        }
    }
}
