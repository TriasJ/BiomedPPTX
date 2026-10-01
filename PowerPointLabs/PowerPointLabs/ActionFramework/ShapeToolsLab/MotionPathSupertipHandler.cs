using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShapeToolsLab
{
    [ExportSupertipRibbonId(ShapeToolsLabText.MotionPathTag)]
    class MotionPathSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return ShapeToolsLabText.MotionPathSupertip;
        }
    }
}
