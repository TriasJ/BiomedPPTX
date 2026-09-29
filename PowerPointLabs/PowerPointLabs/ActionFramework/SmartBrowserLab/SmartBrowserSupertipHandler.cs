using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.SmartBrowserLab
{
    [ExportSupertipRibbonId(SmartBrowserLabText.PaneTag)]
    class SmartBrowserSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return SmartBrowserLabText.ButtonSupertip;
        }
    }
}
