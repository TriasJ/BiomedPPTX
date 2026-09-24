using System;
using System.Collections.Generic;
using System.IO;

using PowerPointLabs.SmartBrowserLab.Models;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SmartBrowserLab.Services
{
    public class ShapeInserter
    {
        private readonly string _smartLibBasePath;
        private readonly Dictionary<string, PowerPoint.Presentation> _presentationCache;

        public ShapeInserter(string smartLibBasePath)
        {
            _smartLibBasePath = smartLibBasePath;
            _presentationCache = new Dictionary<string, PowerPoint.Presentation>(StringComparer.OrdinalIgnoreCase);
        }

        public PowerPoint.Shape InsertAsEditableShape(
            IllustrationItem item,
            PowerPoint.Slide targetSlide,
            PowerPoint.Application app)
        {
            string pptxPath = Path.Combine(_smartLibBasePath, item.PptxFile);
            if (!File.Exists(pptxPath))
            {
                return InsertAsPicture(item, targetSlide, app);
            }

            PowerPoint.Presentation srcPres = GetOrOpenPresentation(pptxPath, app);
            if (srcPres == null)
            {
                return InsertAsPicture(item, targetSlide, app);
            }

            try
            {
                PowerPoint.Slide srcSlide = srcPres.Slides[item.PptxSlide];
                PowerPoint.Shape srcShape = srcSlide.Shapes[item.PptxShapeIndex];

                srcShape.Copy();
                var pastedRange = targetSlide.Shapes.Paste();
                PowerPoint.Shape pasted = pastedRange[1];

                float slideWidth = targetSlide.CustomLayout.Width;
                float slideHeight = targetSlide.CustomLayout.Height;
                pasted.Left = (slideWidth - pasted.Width) / 2;
                pasted.Top = (slideHeight - pasted.Height) / 2;

                return pasted;
            }
            catch
            {
                return InsertAsPicture(item, targetSlide, app);
            }
        }

        public PowerPoint.Shape InsertAsPicture(
            IllustrationItem item,
            PowerPoint.Slide targetSlide,
            PowerPoint.Application app)
        {
            string svgPath = ResolvePath(item.SvgPath);
            string pngPath = ResolvePath(item.PngPath);

            string imagePath = File.Exists(svgPath) ? svgPath : pngPath;
            if (!File.Exists(imagePath))
            {
                return null;
            }

            PowerPoint.Shape shape = targetSlide.Shapes.AddPicture(
                imagePath,
                Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoTrue,
                0, 0);

            float slideWidth = targetSlide.CustomLayout.Width;
            float slideHeight = targetSlide.CustomLayout.Height;
            shape.Left = (slideWidth - shape.Width) / 2;
            shape.Top = (slideHeight - shape.Height) / 2;

            return shape;
        }

        public PowerPoint.Shape InsertTileFromSource(
            IllustrationItem item,
            PowerPoint.Slide targetSlide,
            PowerPoint.Application app)
        {
            return InsertAsEditableShape(item, targetSlide, app);
        }

        public void CloseAllCached()
        {
            foreach (var pres in _presentationCache.Values)
            {
                try
                {
                    pres.Close();
                }
                catch (Exception)
                {
                }
            }
            _presentationCache.Clear();
        }

        private PowerPoint.Presentation GetOrOpenPresentation(string pptxPath, PowerPoint.Application app)
        {
            PowerPoint.Presentation cached;
            if (_presentationCache.TryGetValue(pptxPath, out cached))
            {
                try
                {
                    int count = cached.Slides.Count;
                    return cached;
                }
                catch
                {
                    _presentationCache.Remove(pptxPath);
                }
            }

            try
            {
                PowerPoint.Presentation pres = app.Presentations.Open(
                    pptxPath,
                    ReadOnly: Microsoft.Office.Core.MsoTriState.msoTrue,
                    Untitled: Microsoft.Office.Core.MsoTriState.msoFalse,
                    WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);
                _presentationCache[pptxPath] = pres;
                return pres;
            }
            catch
            {
                return null;
            }
        }

        private string ResolvePath(string relativePath)
        {
            if (string.IsNullOrEmpty(relativePath))
            {
                return "";
            }

            return Path.Combine(_smartLibBasePath, relativePath);
        }
    }
}
