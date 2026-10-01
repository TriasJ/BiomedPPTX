using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShapeToolsLab
{
    [ExportLabelRibbonId(ShapeToolsLabText.GearStarTag)]
    class GearStarLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return ShapeToolsLabText.GearStarLabel;
        }
    }
}
