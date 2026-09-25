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
            if (string.IsNullOrEmpty(item.PptxFile) || item.PptxSlide <= 0)
            {
                return InsertAsSvg(item, targetSlide, app);
            }

            string pptxPath = FindPptxFile(item.PptxFile);
            if (pptxPath == null)
            {
                return InsertAsSvg(item, targetSlide, app);
            }

            PowerPoint.Presentation srcPres = GetOrOpenPresentation(pptxPath, app);
            if (srcPres == null)
            {
                return InsertAsSvg(item, targetSlide, app);
            }

            try
            {
                PowerPoint.Slide srcSlide = srcPres.Slides[item.PptxSlide];
                PowerPoint.Shape srcShape = srcSlide.Shapes[item.PptxShapeIndex];

                return CopyShapeWithRetry(srcShape, targetSlide);
            }
            catch (Exception)
            {
                return InsertAsSvg(item, targetSlide, app);
            }
        }

        public PowerPoint.Shape InsertAsSvg(
            IllustrationItem item,
            PowerPoint.Slide targetSlide,
            PowerPoint.Application app)
        {
            string svgPath = FindAssetFile(item.SvgPath);
            if (svgPath != null)
            {
                return InsertPictureFile(svgPath, targetSlide);
            }

            return InsertAsPng(item, targetSlide, app);
        }

        public PowerPoint.Shape InsertAsPng(
            IllustrationItem item,
            PowerPoint.Slide targetSlide,
            PowerPoint.Application app)
        {
            string pngPath = FindAssetFile(item.PngPath);
            if (pngPath != null)
            {
                return InsertPictureFile(pngPath, targetSlide);
            }

            return null;
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

        private PowerPoint.Shape CopyShapeWithRetry(PowerPoint.Shape srcShape, PowerPoint.Slide targetSlide)
        {
            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    srcShape.Copy();
                    System.Threading.Thread.Sleep(300 + attempt * 200);
                    var pastedRange = targetSlide.Shapes.Paste();
                    System.Threading.Thread.Sleep(100);
                    PowerPoint.Shape pasted = pastedRange[1];

                    float slideWidth = targetSlide.CustomLayout.Width;
                    float slideHeight = targetSlide.CustomLayout.Height;
                    pasted.Left = (slideWidth - pasted.Width) / 2;
                    pasted.Top = (slideHeight - pasted.Height) / 2;

                    return pasted;
                }
                catch (Exception)
                {
                    System.Threading.Thread.Sleep(500);
                }
            }

            return null;
        }

        private string FindPptxFile(string pptxFileName)
        {
            if (string.IsNullOrEmpty(pptxFileName))
            {
                return null;
            }

            string appDataAssets = Path.Combine(ThisAddIn.AppDataFolder, "Assets", "SMART-Lib");
            string installDir = AppDomain.CurrentDomain.BaseDirectory;

            string[] searchPaths = new[]
            {
                Path.Combine(installDir, "Assets", "SMART-Lib", pptxFileName),
                Path.Combine(appDataAssets, pptxFileName),
                Path.Combine(_smartLibBasePath, pptxFileName),
                Path.Combine(_smartLibBasePath, "..", "SMART-Lib", pptxFileName),
                Path.Combine(_smartLibBasePath, "..", pptxFileName),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "__Scratch", "SMART-Lib", pptxFileName),
            };

            foreach (string path in searchPaths)
            {
                string fullPath = Path.GetFullPath(path);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            return null;
        }

        private string FindAssetFile(string relativePath)
        {
            if (string.IsNullOrEmpty(relativePath))
            {
                return null;
            }

            string appDataAssets = Path.Combine(ThisAddIn.AppDataFolder, "Assets", "SMART-Library");
            string installDir = AppDomain.CurrentDomain.BaseDirectory;

            string[] searchPaths = new[]
            {
                Path.Combine(installDir, "Assets", "SMART-Library", relativePath),
                Path.Combine(appDataAssets, relativePath),
                Path.Combine(_smartLibBasePath, relativePath),
                Path.Combine(_smartLibBasePath, "..", "SMART-Library", relativePath),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "__Scratch", "SMART-Library", relativePath),
            };

            foreach (string path in searchPaths)
            {
                string fullPath = Path.GetFullPath(path);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            return null;
        }

        private PowerPoint.Shape InsertPictureFile(string filePath, PowerPoint.Slide targetSlide)
        {
            PowerPoint.Shape shape = targetSlide.Shapes.AddPicture(
                filePath,
                Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoTrue,
                0, 0);

            float slideWidth = targetSlide.CustomLayout.Width;
            float slideHeight = targetSlide.CustomLayout.Height;
            shape.Left = (slideWidth - shape.Width) / 2;
            shape.Top = (slideHeight - shape.Height) / 2;

            return shape;
        }

        private PowerPoint.Presentation GetOrOpenPresentation(string pptxPath, PowerPoint.Application app)
        {
            PowerPoint.Presentation cached;
            if (_presentationCache.TryGetValue(pptxPath, out cached))
            {
                try
                {
                    int count = cached.Slides.Count;
                    if (count > 0)
                    {
                        return cached;
                    }

                    _presentationCache.Remove(pptxPath);
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
                pres.Saved = Microsoft.Office.Core.MsoTriState.msoTrue;
                _presentationCache[pptxPath] = pres;
                return pres;
            }
            catch
            {
                return null;
            }
        }
    }
}
