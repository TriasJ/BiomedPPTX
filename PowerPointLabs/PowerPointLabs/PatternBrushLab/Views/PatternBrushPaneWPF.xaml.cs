using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

using Microsoft.Office.Core;

using PowerPointLabs.PatternBrushLab.Services;
using PowerPointLabs.SmartBrowserLab.Models;
using PowerPointLabs.SmartBrowserLab.Services;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PatternBrushLab.Views
{
    public partial class PatternBrushPaneWPF : UserControl
    {
        public static SmartBrowserLab.Views.IllustrationViewModel PendingCustomPattern { get; set; }

        private SmartDatabase _database;
        private ShapeInserter _shapeInserter;
        private TilingEngine _tilingEngine;
        private PathExtractor _pathExtractor;
        private BrushModeController _brushController;
        private string _basePath;
        private ObservableCollection<PatternViewModel> _patterns;
        private PatternViewModel _selectedPattern;
        private bool _isInitialized;

        public PatternBrushPaneWPF()
        {
            InitializeComponent();
            _patterns = new ObservableCollection<PatternViewModel>();
            _tilingEngine = new TilingEngine();
            _pathExtractor = new PathExtractor();
            _brushController = new BrushModeController();
            patternList.ItemsSource = _patterns;
        }

        public void Initialize(string dbPath, string assetsBasePath)
        {
            if (_isInitialized)
            {
                return;
            }

            try
            {
                _basePath = assetsBasePath;
                _database = new SmartDatabase(dbPath, assetsBasePath);
                _shapeInserter = new ShapeInserter(Path.Combine(assetsBasePath, ".."));
                LoadTileablePatterns();
                _isInitialized = true;

                if (PendingCustomPattern != null)
                {
                    AddCustomPattern(PendingCustomPattern);
                    PendingCustomPattern = null;
                }
            }
            catch (Exception ex)
            {
                selectedPatternText.Text = "Error: " + ex.Message;
            }
        }

        public void ApplyPatternToSelection()
        {
            if (_selectedPattern == null)
            {
                return;
            }

            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;

                if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    string msg = _selectedPattern.Axis == "both"
                        ? "Please select a shape to define the fill area."
                        : "Please select a line or freeform shape first.";
                    System.Windows.MessageBox.Show(msg,
                        "BiomedPPTX", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                PowerPoint.Shape shape = sel.ShapeRange[1];
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null)
                {
                    return;
                }

                var pathPoints = _pathExtractor.ExtractPath(shape);
                float overlapX = (float)offsetXSlider.Value;
                float overlapY = (float)offsetYSlider.Value;

                List<TilePlacement> placements;

                float scatter = (float)scatterSlider.Value;
                float jitter = (float)jitterSlider.Value;
                float angleOffset = (float)angleOffsetSlider.Value;

                string mode = GetApplyMode();
                bool useGridFill;

                if (mode == "fill")
                {
                    useGridFill = true;
                }
                else if (mode == "outline")
                {
                    useGridFill = false;
                }
                else
                {
                    useGridFill = _selectedPattern.Axis == "both";
                }

                if (useGridFill)
                {
                    var bounds = new RectangleF(shape.Left, shape.Top, shape.Width, shape.Height);
                    float offsetPt = _selectedPattern.OffsetAmount > 0
                        ? _selectedPattern.OffsetAmount * (_selectedPattern.TileWidth / _selectedPattern.SvgWidth)
                        : _selectedPattern.TileWidth * 0.5f;

                    placements = _tilingEngine.ComputeGridPlacements(
                        bounds,
                        _selectedPattern.TileWidth,
                        _selectedPattern.TileHeight,
                        overlapX, overlapY,
                        _selectedPattern.OffsetRows,
                        offsetPt);
                }
                else
                {
                    placements = _tilingEngine.ComputePlacements(
                        pathPoints,
                        _selectedPattern.TileWidth,
                        overlapX,
                        scatterPt: scatter,
                        rotationJitterDeg: jitter,
                        rotationOffset: angleOffset);
                }

                if (placements.Count == 0)
                {
                    return;
                }

                var tileItem = new IllustrationItem
                {
                    PptxFile = _selectedPattern.PptxFile,
                    PptxSlide = _selectedPattern.PptxSlide,
                    PptxShapeIndex = _selectedPattern.PptxShapeIndex,
                    Name = _selectedPattern.Name
                };

                PowerPoint.Shape firstTile = _shapeInserter.InsertTileFromSource(tileItem, slide, app);
                if (firstTile == null)
                {
                    return;
                }

                firstTile.Left = placements[0].Position.X - _selectedPattern.TileWidth / 2;
                firstTile.Top = placements[0].Position.Y - _selectedPattern.TileHeight / 2;
                firstTile.Rotation = placements[0].RotationDegrees;

                var tileNames = new List<string> { firstTile.Name };

                for (int i = 1; i < placements.Count; i++)
                {
                    try
                    {
                        PowerPoint.Shape dup = firstTile.Duplicate()[1];
                        dup.Left = placements[i].Position.X - _selectedPattern.TileWidth / 2;
                        dup.Top = placements[i].Position.Y - _selectedPattern.TileHeight / 2;
                        dup.Rotation = placements[i].RotationDegrees;
                        tileNames.Add(dup.Name);
                    }
                    catch (Exception)
                    {
                    }
                }

                if (tileNames.Count > 1)
                {
                    try
                    {
                        var group = slide.Shapes.Range(tileNames.ToArray()).Group();
                        group.Name = "PatternBrush_" + _selectedPattern.Name;
                    }
                    catch (Exception)
                    {
                    }
                }

                try
                {
                    shape.Delete();
                }
                catch (Exception)
                {
                }
                selectedPatternText.Text = string.Format("Applied {0} tiles. Select another shape to apply again.", placements.Count);
                applyButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                selectedPatternText.Text = "Error: " + ex.Message;
                applyButton.IsEnabled = true;
            }
        }

        public void AddCustomPattern(SmartBrowserLab.Views.IllustrationViewModel item)
        {
            string axis = "horizontal";
            if (item.Source == "both")
            {
                axis = "both";
            }

            var pattern = new PatternViewModel
            {
                Id = item.Id,
                Name = item.Name,
                DisplayName = item.DisplayName,
                ThumbnailPath = item.ThumbnailPath,
                PptxFile = item.PptxFile,
                PptxSlide = item.PptxSlide,
                PptxShapeIndex = item.PptxShapeIndex,
                Axis = axis,
                DefaultOverlap = 0,
                TileWidth = item.Width > 0 ? item.Width : 50,
                TileHeight = item.Height > 0 ? item.Height : 50,
                OffsetRows = axis == "both",
                OffsetAmount = 0,
                SvgWidth = item.Width > 0 ? item.Width : 50,
                BorderColor = axis == "both" ? "#FF8C00" : "#9B59B6"
            };

            if (!_patterns.Any(p => p.Id == pattern.Id))
            {
                _patterns.Insert(0, pattern);
            }

            patternList.SelectedItem = pattern;
        }

        #region Event Handlers

        private void PatternBrushPaneWPF_Loaded(object sender, RoutedEventArgs e)
        {
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string axis = GetSelectedAxis();
            string searchText = searchBox.Text != null ? searchBox.Text.Trim() : null;
            LoadTileablePatterns(axis, searchText);
        }

        private void AxisFilter_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized)
            {
                return;
            }

            string axis = GetSelectedAxis();
            string searchText = searchBox != null && searchBox.Text != null ? searchBox.Text.Trim() : null;
            LoadTileablePatterns(axis, searchText);
        }

        private void PatternList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            PatternViewModel newSelection = patternList.SelectedItem as PatternViewModel;
            if (newSelection != null)
            {
                _selectedPattern = newSelection;
                selectedPatternText.Text = string.Format("{0} ({1})", _selectedPattern.Name, _selectedPattern.Axis);
                offsetXSlider.Value = _selectedPattern.DefaultOverlap;
                applyButton.IsEnabled = true;

                string mode = GetApplyMode();
                if (_selectedPattern.Axis == "both" || mode == "fill")
                {
                    offsetYGrid.Visibility = Visibility.Visible;
                    offsetYSlider.Value = _selectedPattern.DefaultOverlap * 0.67;
                }
                else
                {
                    offsetYGrid.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void OffsetXSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (offsetXValueText != null)
            {
                offsetXValueText.Text = string.Format("{0} pt", (int)offsetXSlider.Value);
            }
        }

        private void OffsetYSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (offsetYValueText != null)
            {
                offsetYValueText.Text = string.Format("{0} pt", (int)offsetYSlider.Value);
            }
        }

        private void ApplyMode_Changed(object sender, SelectionChangedEventArgs e)
        {
            string mode = GetApplyMode();
            if (mode == "fill")
            {
                if (offsetYGrid != null)
                {
                    offsetYGrid.Visibility = Visibility.Visible;
                }
            }
            else if (mode == "outline")
            {
                if (offsetYGrid != null)
                {
                    offsetYGrid.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void DrawButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null)
                {
                    return;
                }

                float sw = app.ActivePresentation.PageSetup.SlideWidth;
                float sh = app.ActivePresentation.PageSetup.SlideHeight;
                float margin = 80;

                PowerPoint.Shape line = slide.Shapes.AddLine(
                    margin, sh / 2,
                    sw - margin, sh / 2);

                line.Line.Weight = 2;
                line.Line.ForeColor.RGB = 0xCCCCCC;
                line.Line.DashStyle = MsoLineDashStyle.msoLineDash;
                line.Name = "PatternBrush_Guide";
                line.Select();

                selectedPatternText.Text = "Adjust the line, then Apply. Right-click > Edit Points for curves.";
            }
            catch (Exception ex)
            {
                selectedPatternText.Text = "Error: " + ex.Message;
            }
        }

        private void AngleOffsetSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (angleOffsetValueText != null)
            {
                angleOffsetValueText.Text = (int)angleOffsetSlider.Value + " deg";
            }
        }

        private void ScatterSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (scatterValueText != null)
            {
                scatterValueText.Text = (int)scatterSlider.Value + " pt";
            }
        }

        private void JitterSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (jitterValueText != null)
            {
                jitterValueText.Text = (int)jitterSlider.Value + " deg";
            }
        }

        private void AutoConvert_Changed(object sender, RoutedEventArgs e)
        {
            if (autoConvertToggle.IsChecked == true && _selectedPattern != null)
            {
                _brushController.Activate(OnNewShapeDrawn);
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        Globals.ThisAddIn.Application.CommandBars.ExecuteMso("ShapeScribble");
                    }
                    catch (Exception)
                    {
                        try
                        {
                            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("ShapeFreeform");
                        }
                        catch (Exception)
                        {
                        }
                    }
                }), System.Windows.Threading.DispatcherPriority.Background);
            }
            else
            {
                _brushController.Deactivate();
            }
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            ApplyPatternToSelection();
        }

        #endregion

        #region Private Methods

        private void LoadTileablePatterns(string axisFilter = null, string searchQuery = null)
        {
            if (_database == null)
            {
                return;
            }

            _patterns.Clear();

            string[] tileableTags = { "tileable-1d", "tileable-2d", "tile-horizontal", "tile-vertical" };
            var allItems = new Dictionary<int, IllustrationItem>();

            foreach (string tag in tileableTags)
            {
                var items = _database.GetByTag(tag, 0, 500);
                foreach (var item in items)
                {
                    allItems[item.Id] = item;
                }
            }

            var filtered = allItems.Values.AsEnumerable();

            if (!string.IsNullOrWhiteSpace(searchQuery))
            {
                string q = searchQuery.ToLower();
                filtered = filtered.Where(i =>
                    i.Name.ToLower().Contains(q) ||
                    (i.Description != null && i.Description.ToLower().Contains(q)));
            }

            foreach (var item in filtered.OrderBy(i => i.Name))
            {
                var meta = _database.GetTilingMetadata(item.Id);
                string axis = meta != null ? meta.Axis : "horizontal";
                if (axis == null)
                {
                    axis = "horizontal";
                }

                if (axisFilter != null)
                {
                    if (axisFilter == "horizontal" && axis != "horizontal")
                    {
                        continue;
                    }

                    if (axisFilter == "vertical" && axis != "vertical")
                    {
                        continue;
                    }

                    if (axisFilter == "both" && axis != "both")
                    {
                        continue;
                    }
                }

                string thumbPath = _database.ResolveAssetPath(item.PngPath);
                string displayName = item.Name.Replace("-", " ").Replace("_", " ");
                if (displayName.Length > 20)
                {
                    displayName = displayName.Substring(0, 17) + "...";
                }

                string borderColor = "Transparent";
                if (axis == "horizontal")
                {
                    borderColor = "#4A90D9";
                }
                else if (axis == "vertical")
                {
                    borderColor = "#50C878";
                }
                else if (axis == "both")
                {
                    borderColor = "#FF8C00";
                }

                _patterns.Add(new PatternViewModel
                {
                    Id = item.Id,
                    Name = item.Name,
                    DisplayName = displayName,
                    ThumbnailPath = File.Exists(thumbPath) ? thumbPath : "",
                    PptxFile = item.PptxFile,
                    PptxSlide = item.PptxSlide,
                    PptxShapeIndex = item.PptxShapeIndex,
                    Axis = axis,
                    DefaultOverlap = meta != null && meta.OverlapPt > 0 ? meta.OverlapPt : 10,
                    TileWidth = meta != null && meta.PptxWidth > 0 ? meta.PptxWidth : item.Width,
                    TileHeight = meta != null && meta.PptxHeight > 0 ? meta.PptxHeight : item.Height,
                    OffsetRows = meta != null ? meta.OffsetRows : false,
                    OffsetAmount = meta != null ? meta.OffsetAmount : 0,
                    SvgWidth = meta != null && meta.SvgWidth > 0 ? meta.SvgWidth : item.Width,
                    BorderColor = borderColor
                });
            }
        }

        private string GetApplyMode()
        {
            if (applyModeBox == null || applyModeBox.SelectedItem == null)
            {
                return "auto";
            }

            ComboBoxItem item = applyModeBox.SelectedItem as ComboBoxItem;
            if (item != null && item.Tag != null)
            {
                return item.Tag.ToString();
            }

            return "auto";
        }

        private string GetSelectedAxis()
        {
            if (filterHorizontal != null && filterHorizontal.IsChecked == true)
            {
                return "horizontal";
            }

            if (filterVertical != null && filterVertical.IsChecked == true)
            {
                return "vertical";
            }

            if (filter2D != null && filter2D.IsChecked == true)
            {
                return "both";
            }

            return null;
        }

        private void OnNewShapeDrawn(PowerPoint.Shape shape)
        {
            if (_selectedPattern == null)
            {
                return;
            }

            try
            {
                var app = Globals.ThisAddIn.Application;
                var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null)
                {
                    return;
                }

                var pathPoints = _pathExtractor.ExtractPath(shape);
                float overlap = (float)offsetXSlider.Value;
                float scatter = (float)scatterSlider.Value;
                float jitter = (float)jitterSlider.Value;
                float angleOffset = (float)angleOffsetSlider.Value;

                var placements = _tilingEngine.ComputePlacements(
                    pathPoints, _selectedPattern.TileWidth, overlap,
                    scatterPt: scatter, rotationJitterDeg: jitter,
                    rotationOffset: angleOffset);

                if (placements.Count == 0)
                {
                    return;
                }

                var tileItem = new IllustrationItem
                {
                    PptxFile = _selectedPattern.PptxFile,
                    PptxSlide = _selectedPattern.PptxSlide,
                    PptxShapeIndex = _selectedPattern.PptxShapeIndex,
                    Name = _selectedPattern.Name
                };

                PowerPoint.Shape firstTile = _shapeInserter.InsertTileFromSource(tileItem, slide, app);
                if (firstTile == null)
                {
                    return;
                }

                firstTile.Left = placements[0].Position.X - _selectedPattern.TileWidth / 2;
                firstTile.Top = placements[0].Position.Y - _selectedPattern.TileHeight / 2;
                firstTile.Rotation = placements[0].RotationDegrees;

                var tileNames = new List<string> { firstTile.Name };
                for (int i = 1; i < placements.Count; i++)
                {
                    PowerPoint.Shape dup = firstTile.Duplicate()[1];
                    dup.Left = placements[i].Position.X - _selectedPattern.TileWidth / 2;
                    dup.Top = placements[i].Position.Y - _selectedPattern.TileHeight / 2;
                    dup.Rotation = placements[i].RotationDegrees;
                    tileNames.Add(dup.Name);
                }

                if (tileNames.Count > 1)
                {
                    slide.Shapes.Range(tileNames.ToArray()).Group().Name =
                        "PatternBrush_" + _selectedPattern.Name;
                }

                shape.Delete();
                selectedPatternText.Text = string.Format("Auto-applied {0} tiles", placements.Count);
            }
            catch (Exception ex)
            {
                selectedPatternText.Text = "Auto-convert error: " + ex.Message;
            }
        }

        #endregion
    }
}
