using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShapeToolsLab
{
    [ExportSupertipRibbonId(ShapeToolsLabText.WavyLineTag)]
    class WavyLineSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return ShapeToolsLabText.WavyLineSupertip;
        }
    }
}
