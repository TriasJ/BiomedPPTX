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
                float overlap = (float)overlapSlider.Value;

                List<TilePlacement> placements;

                float scatter = (float)scatterSlider.Value;
                float jitter = (float)jitterSlider.Value;
                float angleOffset = (float)angleOffsetSlider.Value;

                if (_selectedPattern.Axis == "both")
                {
                    var bounds = new RectangleF(shape.Left, shape.Top, shape.Width, shape.Height);
                    float overlapX = overlap > 0 ? overlap : 18f;
                    float overlapY = overlap > 0 ? overlap * 0.67f : 12f;
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

                    shape.Delete();
                }
                else
                {
                    placements = _tilingEngine.ComputePlacements(
                        pathPoints,
                        _selectedPattern.TileWidth,
                        overlap,
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
                    PowerPoint.Shape dup = firstTile.Duplicate()[1];
                    dup.Left = placements[i].Position.X - _selectedPattern.TileWidth / 2;
                    dup.Top = placements[i].Position.Y - _selectedPattern.TileHeight / 2;
                    dup.Rotation = placements[i].RotationDegrees;
                    tileNames.Add(dup.Name);
                }

                if (tileNames.Count > 1)
                {
                    var group = slide.Shapes.Range(tileNames.ToArray()).Group();
                    group.Name = "PatternBrush_" + _selectedPattern.Name;
                }

                shape.Delete();
                selectedPatternText.Text = string.Format("Applied {0} tiles", placements.Count);
            }
            catch (Exception ex)
            {
                selectedPatternText.Text = "Error: " + ex.Message;
            }
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
            _selectedPattern = patternList.SelectedItem as PatternViewModel;
            if (_selectedPattern != null)
            {
                selectedPatternText.Text = string.Format("{0} ({1})", _selectedPattern.Name, _selectedPattern.Axis);
                overlapSlider.Value = _selectedPattern.DefaultOverlap;
                applyButton.IsEnabled = true;
            }
            else
            {
                selectedPatternText.Text = "No pattern selected";
                applyButton.IsEnabled = false;
            }
        }

        private void OverlapSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (overlapValueText != null)
            {
                overlapValueText.Text = (int)overlapSlider.Value + " pt";
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
                float overlap = (float)overlapSlider.Value;
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
