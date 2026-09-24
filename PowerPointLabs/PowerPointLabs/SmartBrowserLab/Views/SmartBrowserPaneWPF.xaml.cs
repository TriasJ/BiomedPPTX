using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
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
                    var bioArtResults = _bioArtFetcher.Search(searchQuery, 30);
                    string bioArtCache = Path.Combine(ThisAddIn.AppDataFolder, "BioArtCache");
                    foreach (var ba in bioArtResults)
                    {
                        string displayName = ba.Title;
                        if (displayName.Length > 30)
                        {
                            displayName = displayName.Substring(0, 27) + "...";
                        }

                        string bioArtThumb = "";
                        string[] exts = { "svg", "png", "jpg" };
                        foreach (string ext in exts)
                        {
                            string cached = Path.Combine(bioArtCache, string.Format("bioart_{0}.{1}", ba.Id, ext));
                            if (File.Exists(cached))
                            {
                                bioArtThumb = cached;
                                break;
                            }
                        }

                        _illustrations.Add(new IllustrationViewModel
                        {
                            Id = ba.Id + 100000,
                            Name = ba.Title,
                            DisplayName = string.Format("[BioArt] {0}", displayName),
                            Description = string.Format("{0} [BioArt - {1}]", ba.Description, ba.License),
                            ThumbnailPath = bioArtThumb,
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
                    string imagePath = selected.ThumbnailPath;
                    if (string.IsNullOrEmpty(imagePath) || !File.Exists(imagePath))
                    {
                        statusText.Text = "Downloading BioArt image...";
                        if (_bioArtFetcher != null)
                        {
                            int bioArtId = selected.Id - 100000;
                            var bioArtResults = _bioArtFetcher.Search(selected.Name, 1);
                            if (bioArtResults.Count > 0)
                            {
                                imagePath = _bioArtFetcher.DownloadImage(bioArtResults[0]);
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
                        statusText.Text = string.Format("Inserted BioArt: {0}", selected.Name);
                    }
                    else
                    {
                        statusText.Text = "Failed to download BioArt image";
                    }
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
                }

                statusText.Text = "Inserted: " + selected.Name;
            }
            catch (Exception ex)
            {
                statusText.Text = "Insert error: " + ex.Message;
            }
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

        private void UseAsPattern_Click(object sender, RoutedEventArgs e)
        {
            var selected = illustrationList.SelectedItem as IllustrationViewModel;
            if (selected == null)
            {
                return;
            }

            // TODO: Open Pattern Brush pane with this illustration pre-selected
            statusText.Text = "Pattern brush: " + selected.Name + " (coming soon)";
        }

        #endregion
    }
}
