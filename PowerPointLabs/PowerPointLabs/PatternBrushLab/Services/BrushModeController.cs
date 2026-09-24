using System;
using System.Collections.Generic;

using Microsoft.Office.Core;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PatternBrushLab.Services
{
    public class BrushModeController
    {
        private bool _isActive;
        private HashSet<string> _knownShapeNames;
        private Action<PowerPoint.Shape> _onNewShapeDetected;

        public bool IsActive
        {
            get { return _isActive; }
        }

        public void Activate(Action<PowerPoint.Shape> onNewShapeDetected)
        {
            if (_isActive)
            {
                return;
            }

            _onNewShapeDetected = onNewShapeDetected;
            _isActive = true;
            _knownShapeNames = SnapshotCurrentShapes();

            Globals.ThisAddIn.Application.WindowSelectionChange += OnSelectionChange;
        }

        public void Deactivate()
        {
            if (!_isActive)
            {
                return;
            }

            _isActive = false;
            try
            {
                Globals.ThisAddIn.Application.WindowSelectionChange -= OnSelectionChange;
            }
            catch (Exception)
            {
            }

            _onNewShapeDetected = null;
            _knownShapeNames = null;
        }

        private void OnSelectionChange(PowerPoint.Selection sel)
        {
            if (!_isActive || _onNewShapeDetected == null)
            {
                return;
            }

            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }

            try
            {
                PowerPoint.Shape shape = sel.ShapeRange[1];

                if (_knownShapeNames.Contains(shape.Name))
                {
                    return;
                }

                if (shape.Type == MsoShapeType.msoFreeform ||
                    shape.Type == MsoShapeType.msoLine ||
                    (shape.Type == MsoShapeType.msoAutoShape &&
                     shape.AutoShapeType == MsoAutoShapeType.msoShapeMixed))
                {
                    _knownShapeNames.Add(shape.Name);
                    _onNewShapeDetected(shape);
                    _knownShapeNames = SnapshotCurrentShapes();
                }
            }
            catch (Exception)
            {
            }
        }

        private HashSet<string> SnapshotCurrentShapes()
        {
            var names = new HashSet<string>();
            try
            {
                var slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide != null)
                {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        names.Add(shape.Name);
                    }
                }
            }
            catch (Exception)
            {
            }

            return names;
        }
    }
}
