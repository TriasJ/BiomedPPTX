using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShapeToolsLab
{
    [ExportActionRibbonId(ShapeToolsLabText.WavyLineTag)]
    class WavyLineActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var dialog = new PowerPointLabs.ShapeToolsLab.Views.WavyLineDialog();
            dialog.ShowDialog();
        }
    }
}
