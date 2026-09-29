using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PatternBrushLab
{
    [ExportSupertipRibbonId(PatternBrushLabText.PaneTag)]
    class PatternBrushSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return PatternBrushLabText.ButtonSupertip;
        }
    }
}
