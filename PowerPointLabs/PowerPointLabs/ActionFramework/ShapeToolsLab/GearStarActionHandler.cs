using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShapeToolsLab
{
    [ExportActionRibbonId(ShapeToolsLabText.GearStarTag)]
    class GearStarActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var dialog = new PowerPointLabs.ShapeToolsLab.Views.GearStarDialog();
            dialog.ShowDialog();
        }
    }
}
