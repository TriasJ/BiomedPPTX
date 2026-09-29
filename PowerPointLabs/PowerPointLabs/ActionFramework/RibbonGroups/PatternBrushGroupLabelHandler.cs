using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.RibbonGroups
{
    [ExportLabelRibbonId(CommonText.PatternBrushGroupId)]
    class PatternBrushGroupLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return CommonText.PatternBrushGroupLabel;
        }
    }
}
