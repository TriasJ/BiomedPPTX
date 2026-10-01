using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShapeToolsLab
{
    [ExportLabelRibbonId(ShapeToolsLabText.MotionPathTag)]
    class MotionPathLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return ShapeToolsLabText.MotionPathLabel;
        }
    }
}
