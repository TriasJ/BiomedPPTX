using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PatternBrushLab
{
    [ExportLabelRibbonId(PatternBrushLabText.PaneTag)]
    class PatternBrushLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return PatternBrushLabText.ButtonLabel;
        }
    }
}
