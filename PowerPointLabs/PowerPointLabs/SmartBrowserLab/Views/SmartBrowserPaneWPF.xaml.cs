using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;

using PowerPointLabs.SmartBrowserLab.Models;
using PowerPointLabs.SmartBrowserLab.Services;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SmartBrowserLab.Views
{
    public partial class SmartBrowserPaneWPF : UserControl
    {
        private SmartDatabase _database;
        private BioArtFetcher _bioArtFetcher;
        private ShapeInserter _shapeInserter;
        private string _basePath;
        private DispatcherTimer _searchDebounce;
        private ObservableCollection<IllustrationViewModel> _illustrations;
        private List<string> _activeTagFilters;
        private bool _isInitialized;

        public SmartBrowserPaneWPF()
        {
            InitializeComponent();
            _illustrations = new ObservableCollection<IllustrationViewModel>();
            _activeTagFilters = new List<string>();
            _searchDebounce = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            _searchDebounce.Tick += SearchDebounce_Tick;
            illustrationList.ItemsSource = _illustrations;
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

                string bioArtIndex = FindBioArtIndex(assetsBasePath);
                string bioArtCache = Path.Combine(ThisAddIn.AppDataFolder, "BioArtCache");
                _bioArtFetcher = new BioArtFetcher(bioArtIndex, bioArtCache);

                LoadCategories();
                LoadTagFilters();
                LoadIllustrations();
                _isInitialized = true;
            }
            catch (Exception ex)
            {
                statusText.Text = "Error: " + ex.Message;
            }
        }

        private void LoadCategories()
        {
            if (_database == null)
            {
                return;
            }

            categoryBox.Items.Clear();
            categoryBox.Items.Add(new ComboBoxItem { Content = "All Categories", Tag = -1 });

            foreach (var topic in _database.GetTopics())
            {
                categoryBox.Items.Add(new ComboBoxItem
                {
                    Content = string.Format("{0} ({1})", topic.Name, topic.IllustrationCount),
                    Tag = topic.Id
                });
            }

            categoryBox.SelectedIndex = 0;
        }

        private void LoadTagFilters()
        {
            if (_database == null)
            {
                return;
            }

            tagPanel.Children.Clear();
            foreach (var tag in _database.GetTags())
            {
                var toggle = new ToggleButton
                {
                    Content = string.Format("{0} ({1})", tag.Name, tag.Count),
                    Tag = tag.Name,
                    Style = FindResource("TagToggleStyle") as Style
                };
                toggle.Checked += TagToggle_Changed;
                toggle.Unchecked += TagToggle_Changed;
                tagPanel.Children.Add(toggle);
            }
        }

        private void LoadIllustrations(string searchQuery = null, int topicId = -1)
        {
            if (_database == null)
            {
                return;
            }

            try
            {
                List<IllustrationItem> items;

                if (!string.IsNullOrWhiteSpace(searchQuery))
                {
                    items = _database.Search(searchQuery, 200);
                }
                else if (topicId > 0)
                {
                    items = _database.GetByTopic(topicId, 0, 200);
                }
                else if (_activeTagFilters.Count > 0)
                {
                    items = _database.GetByTag(_activeTagFilters[0], 0, 500);
                    foreach (var tag in _activeTagFilters.Skip(1))
                    {
                        var tagItems = _database.GetByTag(tag, 0, 500);
                        var tagIds = new HashSet<int>(tagItems.Select(t => t.Id));
                        items = items.Where(i => tagIds.Contains(i.Id)).ToList();
                    }
                }
                else
                {
                    items = _database.GetByTopic(-1, 0, 200);
                    if (items.Count == 0)
                    {
                        items = _database.Search("*", 200);
                    }
                }

                _illustrations.Clear();
                foreach (var item in items)
                {
                    string thumbPath = _database.ResolveAssetPath(item.PngPath);
                    string displayName = item.Name.Replace("-", " ").Replace("_", " ");
                    if (displayName.Length > 30)
                    {
                        displayName = displayName.Substring(0, 27) + "...";
                    }

                    _illustrations.Add(new IllustrationViewModel
                    {
                        Id = item.Id,
                        Name = item.Name,
                        DisplayName = displayName,
                        Description = item.Description,
                        ThumbnailPath = File.Exists(thumbPath) ? thumbPath : "",
                        SvgPath = item.SvgPath,
                        PptxFile = item.PptxFile,
                        PptxSlide = item.PptxSlide,
                        PptxShapeIndex = item.PptxShapeIndex,
                        Width = item.Width,
                        Height = item.Height,
                        Topic = item.Topic,
                        Source = "SMART"
                    });
                }

                if (bioArtToggle.IsChecked == true && !string.IsNullOrWhiteSpace(searchQuery) && _bioArtFetcher != null)
                {
                    var bioArtResults = _bioArtFetcher.Search(searchQuery, 20);
                    foreach (var ba in bioArtResults)
                    {
                        string displayName = ba.Title;
                        if (displayName.Length > 25)
                        {
                            displayName = displayName.Substring(0, 22) + "...";
                        }

                        string cachedThumb = _bioArtFetcher.GetCachedPath(ba.Id);

                        _illustrations.Add(new IllustrationViewModel
                        {
                            Id = ba.Id + 100000,
                            Name = ba.Title,
                            DisplayName = cachedThumb != null ? displayName : "[BioArt] " + displayName,
                            Description = string.Format("{0} [BioArt - {1}]", ba.Description, ba.License),
                            ThumbnailPath = cachedThumb != null ? cachedThumb : "",
                            SvgPath = "",
                            PptxFile = "",
                            PptxSlide = 0,
                            PptxShapeIndex = 0,
                            Width = 0,
                            Height = 0,
                            Topic = "BioArt",
                            Source = "BioArt"
                        });
                    }
                }

                statusText.Text = string.Format("Showing {0} illustrations", _illustrations.Count);
            }
            catch (Exception ex)
            {
                statusText.Text = "Search error: " + ex.Message;
            }
        }

        private string FindBioArtIndex(string assetsBasePath)
        {
            string[] candidates = new[]
            {
                Path.Combine(assetsBasePath, "bioart_index.json"),
                Path.Combine(ThisAddIn.AppDataFolder, "Assets", "bioart_index.json"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                    ".claude", "skills", "fetch-media", "bioart_index.json")
            };

            foreach (string path in candidates)
            {
                if (File.Exists(path))
                {
                    return path;
                }
            }

            return candidates[0];
        }

        private void InsertSelectedShape(bool asEditable)
        {
            var selected = illustrationList.SelectedItem as IllustrationViewModel;
            if (selected == null)
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

                if (selected.Source == "BioArt")
                {
                    InsertBioArtShape(selected, slide);
                }
                else
                {
                    var item = new IllustrationItem
                    {
                        Id = selected.Id,
                        Name = selected.Name,
                        PptxFile = selected.PptxFile,
                        PptxSlide = selected.PptxSlide,
                        PptxShapeIndex = selected.PptxShapeIndex,
                        SvgPath = selected.SvgPath,
                        PngPath = selected.ThumbnailPath
                    };

                    if (asEditable)
                    {
                        _shapeInserter.InsertAsEditableShape(item, slide, app);
                    }
                    else
                    {
                        _shapeInserter.InsertAsSvg(item, slide, app);
                    }

                    statusText.Text = "Inserted: " + selected.Name;
                }
            }
            catch (Exception ex)
            {
                statusText.Text = "Insert error: " + ex.Message;
            }
        }

        private void InsertBioArtShape(IllustrationViewModel selected, PowerPoint.Slide slide)
        {
            string imagePath = selected.ThumbnailPath;

            if (string.IsNullOrEmpty(imagePath) || !File.Exists(imagePath))
            {
                statusText.Text = "Acquiring BioArt: " + selected.Name + "...";

                if (_bioArtFetcher != null)
                {
                    int bioArtId = selected.Id - 100000;
                    BioArtItem bioItem = _bioArtFetcher.FindById(bioArtId);
                    if (bioItem != null)
                    {
                        imagePath = _bioArtFetcher.DownloadImage(bioItem);
                    }
                }
            }

            if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
            {
                PowerPoint.Shape bioShape = slide.Shapes.AddPicture(
                    imagePath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    0, 0);
                float sw = slide.CustomLayout.Width;
                float sh = slide.CustomLayout.Height;
                bioShape.Left = (sw - bioShape.Width) / 2;
                bioShape.Top = (sh - bioShape.Height) / 2;

                selected.ThumbnailPath = imagePath;
                statusText.Text = "Inserted BioArt: " + selected.Name;
            }
            else
            {
                statusText.Text = "Failed to download BioArt image";
            }
        }

        private void DownloadBioArtThumbnailAsync(IllustrationViewModel viewModel)
        {
            if (_bioArtFetcher == null || _bioArtFetcher.IsDownloading)
            {
                return;
            }

            int bioArtId = viewModel.Id - 100000;
            BioArtItem bioItem = _bioArtFetcher.FindById(bioArtId);
            if (bioItem == null)
            {
                return;
            }

            string itemName = viewModel.Name;
            statusText.Text = "Acquiring: " + itemName + "...";

            Task.Run(() =>
            {
                string path = _bioArtFetcher.DownloadImage(bioItem, "png");
                Dispatcher.Invoke(() =>
                {
                    if (path != null)
                    {
                        viewModel.ThumbnailPath = path;
                        viewModel.DisplayName = viewModel.Name;
                        if (viewModel.DisplayName.Length > 30)
                        {
                            viewModel.DisplayName = viewModel.DisplayName.Substring(0, 27) + "...";
                        }

                        viewModel.OnPropertyChanged("ThumbnailPath");
                        viewModel.OnPropertyChanged("DisplayName");
                        statusText.Text = "Cached: " + itemName;
                    }
                    else
                    {
                        statusText.Text = "Download failed: " + itemName;
                    }
                });
            });
        }

        #region Event Handlers

        private void SmartBrowserPaneWPF_Loaded(object sender, RoutedEventArgs e)
        {
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            searchPlaceholder.Visibility = string.IsNullOrEmpty(searchBox.Text)
                ? Visibility.Visible : Visibility.Collapsed;

            _searchDebounce.Stop();
            _searchDebounce.Start();
        }

        private void SearchDebounce_Tick(object sender, EventArgs e)
        {
            _searchDebounce.Stop();
            string query = searchBox.Text != null ? searchBox.Text.Trim() : null;

            int topicId = -1;
            ComboBoxItem selectedCat = categoryBox.SelectedItem as ComboBoxItem;
            if (selectedCat != null)
            {
                topicId = (int)selectedCat.Tag;
            }

            if (!string.IsNullOrEmpty(query))
            {
                LoadIllustrations(searchQuery: query);
            }
            else
            {
                LoadIllustrations(topicId: topicId);
            }
        }

        private void CategoryBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitialized)
            {
                return;
            }

            int topicId = -1;
            ComboBoxItem selectedCat2 = categoryBox.SelectedItem as ComboBoxItem;
            if (selectedCat2 != null)
            {
                topicId = (int)selectedCat2.Tag;
            }

            if (string.IsNullOrEmpty(searchBox.Text))
            {
                LoadIllustrations(topicId: topicId);
            }
        }

        private void TagToggle_Changed(object sender, RoutedEventArgs e)
        {
            _activeTagFilters.Clear();
            foreach (var child in tagPanel.Children)
            {
                ToggleButton toggle = child as ToggleButton;
                if (toggle != null && toggle.IsChecked == true)
                {
                    _activeTagFilters.Add(toggle.Tag as string);
                }
            }

            if (string.IsNullOrEmpty(searchBox.Text))
            {
                LoadIllustrations();
            }
        }

        private void BioArtToggle_Changed(object sender, RoutedEventArgs e)
        {
            if (_bioArtFetcher != null)
            {
                statusText.Text = bioArtToggle.IsChecked == true
                    ? string.Format("BioArt online enabled ({0} items)", _bioArtFetcher.IndexCount)
                    : "BioArt offline";
            }

            if (!string.IsNullOrEmpty(searchBox.Text))
            {
                _searchDebounce.Stop();
                _searchDebounce.Start();
            }
        }

        private void IllustrationList_DoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            InsertSelectedShape(asEditable: true);
        }

        private void IllustrationList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            insertButton.IsEnabled = illustrationList.SelectedItem != null;

            var selected = illustrationList.SelectedItem as IllustrationViewModel;
            if (selected != null && selected.Source == "BioArt"
                && string.IsNullOrEmpty(selected.ThumbnailPath))
            {
                DownloadBioArtThumbnailAsync(selected);
            }
        }

        private void InsertButton_Click(object sender, RoutedEventArgs e)
        {
            InsertSelectedShape(asEditable: true);
        }

        private void InsertEditable_Click(object sender, RoutedEventArgs e)
        {
            InsertSelectedShape(asEditable: true);
        }

        private void InsertImage_Click(object sender, RoutedEventArgs e)
        {
            InsertSelectedShape(asEditable: false);
        }

        private void UseAsPathPattern_Click(object sender, RoutedEventArgs e)
        {
            SendToPatternBrush("horizontal");
        }

        private void UseAsFillPattern_Click(object sender, RoutedEventArgs e)
        {
            SendToPatternBrush("both");
        }

        private void SendToPatternBrush(string axis)
        {
            var selected = illustrationList.SelectedItem as IllustrationViewModel;
            if (selected == null)
            {
                return;
            }

            selected.Source = axis;
            PatternBrushLab.Views.PatternBrushPaneWPF.PendingCustomPattern = selected;

            try
            {
                Microsoft.Office.Tools.CustomTaskPane brushPane = Globals.ThisAddIn.GetActivePane(typeof(PatternBrushLab.PatternBrushPane));

                if (brushPane == null)
                {
                    var control = new PatternBrushLab.PatternBrushPane();
                    PowerPoint.DocumentWindow wnd = Globals.ThisAddIn.Application.ActiveWindow;
                    brushPane = Globals.ThisAddIn.RegisterTaskPane(
                        control,
                        TextCollection.PatternBrushLabText.TaskPanelTitle,
                        wnd);
                }

                if (brushPane != null)
                {
                    brushPane.Visible = true;
                    PatternBrushLab.PatternBrushPane pane = brushPane.Control as PatternBrushLab.PatternBrushPane;
                    if (pane != null)
                    {
                        string assetsPath = FindSmartLibraryPath();
                        string dbPath = Path.Combine(assetsPath, "illustrations.db");
                        if (File.Exists(dbPath))
                        {
                            pane.InitBrush(dbPath, assetsPath);
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            string modeLabel = axis == "both" ? "Fill (2D)" : "Path (1D)";
            statusText.Text = string.Format("Sent to Pattern Brush as {0}: {1}", modeLabel, selected.Name);
        }

        private string FindSmartLibraryPath()
        {
            string[] searchPaths = new[]
            {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "SMART-Library"),
                Path.Combine(ThisAddIn.AppDataFolder, "Assets", "SMART-Library"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "__Scratch", "SMART-Library")
            };

            foreach (string path in searchPaths)
            {
                if (File.Exists(Path.Combine(path, "illustrations.db")))
                {
                    return path;
                }
            }

            return searchPaths[0];
        }

        private void RecolorToggle_Changed(object sender, RoutedEventArgs e)
        {
            if (recolorPanel != null)
            {
                recolorPanel.Visibility = recolorToggle.IsChecked == true
                    ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private void HueSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (hueValueText != null)
            {
                hueValueText.Text = string.Format("{0}", (int)hueSlider.Value);
            }
        }

        private void SatSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (satValueText != null)
            {
                satValueText.Text = string.Format("{0}%", (int)satSlider.Value);
            }
        }

        private void ApplyRecolor_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;
                if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    statusText.Text = "Select a grouped shape on the slide first";
                    return;
                }

                float hueShift = (float)hueSlider.Value;
                float satScale = (float)satSlider.Value / 100f;

                for (int i = 1; i <= sel.ShapeRange.Count; i++)
                {
                    RecolorShape(sel.ShapeRange[i], hueShift, satScale);
                }

                statusText.Text = string.Format("Recolored with hue {0}, sat {1}%", (int)hueShift, (int)(satScale * 100));
            }
            catch (Exception ex)
            {
                statusText.Text = "Recolor error: " + ex.Message;
            }
        }

        private void RecolorShape(PowerPoint.Shape shape, float hueShift, float satScale)
        {
            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
            {
                for (int i = 1; i <= shape.GroupItems.Count; i++)
                {
                    RecolorShape(shape.GroupItems[i], hueShift, satScale);
                }

                return;
            }

            try
            {
                if (shape.Fill.Visible == Microsoft.Office.Core.MsoTriState.msoTrue &&
                    shape.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillSolid)
                {
                    int rgb = shape.Fill.ForeColor.RGB;
                    int newRgb = ShiftHueSaturation(rgb, hueShift, satScale);
                    if (newRgb != rgb)
                    {
                        shape.Fill.ForeColor.RGB = newRgb;
                    }
                }
            }
            catch (Exception)
            {
            }

            try
            {
                if (shape.Line.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    int rgb = shape.Line.ForeColor.RGB;
                    int newRgb = ShiftHueSaturation(rgb, hueShift, satScale);
                    if (newRgb != rgb)
                    {
                        shape.Line.ForeColor.RGB = newRgb;
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private int ShiftHueSaturation(int rgb, float hueShift, float satScale)
        {
            int r = rgb & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = (rgb >> 16) & 0xFF;

            if (r > 240 && g > 240 && b > 240)
            {
                return rgb;
            }

            if (r < 15 && g < 15 && b < 15)
            {
                return rgb;
            }

            int maxC = Math.Max(r, Math.Max(g, b));
            int minC = Math.Min(r, Math.Min(g, b));
            if (maxC - minC < 20)
            {
                return rgb;
            }

            float rf = r / 255f;
            float gf = g / 255f;
            float bf = b / 255f;
            float max = Math.Max(rf, Math.Max(gf, bf));
            float min = Math.Min(rf, Math.Min(gf, bf));
            float l = (max + min) / 2f;
            float h = 0;
            float s = 0;

            if (max != min)
            {
                float d = max - min;
                s = l > 0.5f ? d / (2f - max - min) : d / (max + min);

                if (max == rf)
                {
                    h = (gf - bf) / d + (gf < bf ? 6f : 0f);
                }
                else if (max == gf)
                {
                    h = (bf - rf) / d + 2f;
                }
                else
                {
                    h = (rf - gf) / d + 4f;
                }

                h /= 6f;
            }

            h += hueShift / 360f;
            while (h < 0)
            {
                h += 1f;
            }

            while (h > 1)
            {
                h -= 1f;
            }

            s = Math.Min(1f, Math.Max(0f, s * satScale));

            float r2;
            float g2;
            float b2;
            if (s == 0)
            {
                r2 = l;
                g2 = l;
                b2 = l;
            }
            else
            {
                float q = l < 0.5f ? l * (1f + s) : l + s - l * s;
                float p = 2f * l - q;
                r2 = HueToRgb(p, q, h + 1f / 3f);
                g2 = HueToRgb(p, q, h);
                b2 = HueToRgb(p, q, h - 1f / 3f);
            }

            int nr = (int)Math.Round(r2 * 255);
            int ng = (int)Math.Round(g2 * 255);
            int nb = (int)Math.Round(b2 * 255);
            nr = Math.Max(0, Math.Min(255, nr));
            ng = Math.Max(0, Math.Min(255, ng));
            nb = Math.Max(0, Math.Min(255, nb));

            return nr | (ng << 8) | (nb << 16);
        }

        private float HueToRgb(float p, float q, float t)
        {
            if (t < 0f)
            {
                t += 1f;
            }

            if (t > 1f)
            {
                t -= 1f;
            }

            if (t < 1f / 6f)
            {
                return p + (q - p) * 6f * t;
            }

            if (t < 1f / 2f)
            {
                return q;
            }

            if (t < 2f / 3f)
            {
                return p + (q - p) * (2f / 3f - t) * 6f;
            }

            return p;
        }

        #endregion
    }
}
