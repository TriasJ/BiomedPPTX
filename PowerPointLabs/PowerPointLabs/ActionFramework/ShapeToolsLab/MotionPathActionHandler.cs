using System;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ShapeToolsLab.Services;
using PowerPointLabs.TextCollection;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ActionFramework.ShapeToolsLab
{
    [ExportActionRibbonId(ShapeToolsLabText.MotionPathTag)]
    class MotionPathActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count < 2)
                {
                    MessageBox.Show(
                        "Select exactly 2 shapes:\n1) The object to animate\n2) The line or freeform path",
                        "BiomedPPTX", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                PowerPoint.Shape target = sel.ShapeRange[1];
                PowerPoint.Shape pathShape = sel.ShapeRange[2];
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;

                if (slide == null)
                {
                    return;
                }

                MotionPathConverter.Convert(target, pathShape, slide);
                MessageBox.Show("Motion path animation created.", "BiomedPPTX",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "BiomedPPTX",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
