using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Threading;

using PowerPointLabs.SmartBrowserLab.Models;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SmartBrowserLab.Views
{
    public partial class MoleculeViewerControl : UserControl
    {
        private PubChemClient _pubChemClient;
        private MoleculeCache _cache;
        private DispatcherTimer _searchDebounce;
        private ObservableCollection<MoleculeViewModel> _molecules;
        private MoleculeViewModel _selectedMolecule;
        private PubChemCompoundData _current3DCompound;
        private PubChemCompoundData _current2DCompound;
        private bool _isInitialized;
        private bool _isSearching;

        private Point _lastMousePos;
        private bool _isDragging;
        private double _cameraDistance = 20.0;
        private double _cameraTheta = 0.0;
        private double _cameraPhi = Math.PI / 4.0;
        private Point3D _cameraTarget = new Point3D(0, 0, 0);

        public MoleculeViewerControl()
        {
            InitializeComponent();
            _molecules = new ObservableCollection<MoleculeViewModel>();
            _searchDebounce = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            _searchDebounce.Tick += SearchDebounce_Tick;
            moleculeList.ItemsSource = _molecules;
        }

        public void EnsureInitialized()
        {
            if (_isInitialized)
            {
                return;
            }

            string cachePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "BiomedPPTX",
                "MoleculeCache");

            _pubChemClient = new PubChemClient();
            _cache = new MoleculeCache(cachePath);
            _isInitialized = true;
        }

        #region Search

        private void MolSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (molSearchPlaceholder != null)
            {
                molSearchPlaceholder.Visibility = string.IsNullOrEmpty(molSearchBox.Text)
                    ? Visibility.Visible
                    : Visibility.Collapsed;
            }

            _searchDebounce.Stop();
            _searchDebounce.Start();
        }

        private void MolSearchBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                _searchDebounce.Stop();
                PerformSearch();
            }
        }

        private void SearchDebounce_Tick(object sender, EventArgs e)
        {
            _searchDebounce.Stop();
            PerformSearch();
        }

        private void PerformSearch()
        {
            string query = molSearchBox.Text;
            if (string.IsNullOrWhiteSpace(query) || query.Length < 2)
            {
                return;
            }

            if (_isSearching)
            {
                return;
            }

            EnsureInitialized();
            _isSearching = true;
            molStatusText.Text = string.Format("Searching for \"{0}\"...", query);

            Task.Run(() =>
            {
                try
                {
                    List<int> cids = _pubChemClient.SearchByName(query);
                    if (cids == null || cids.Count == 0)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            _molecules.Clear();
                            moleculeList.Visibility = Visibility.Collapsed;
                            molStatusText.Text = "No results found.";
                            _isSearching = false;
                        });
                        return;
                    }

                    int maxResults = Math.Min(cids.Count, 10);
                    List<MoleculeViewModel> results = new List<MoleculeViewModel>();

                    for (int i = 0; i < maxResults; i++)
                    {
                        int cid = cids[i];
                        MoleculeSearchResult props = _pubChemClient.GetProperties(cid);

                        MoleculeViewModel vm = new MoleculeViewModel();
                        vm.Cid = cid;

                        if (props != null)
                        {
                            vm.Name = props.Name;
                            vm.Formula = props.MolecularFormula;
                            vm.Weight = string.Format("{0:F2} g/mol", props.MolecularWeight);
                            vm.IupacName = props.IupacName;
                        }

                        string cachedImage = _cache.GetCached2DImage(cid);
                        if (cachedImage == null)
                        {
                            byte[] png = _pubChemClient.Download2DImage(cid);
                            if (png != null)
                            {
                                _cache.Save2DImage(cid, png);
                                cachedImage = _cache.GetCached2DImage(cid);
                            }
                        }

                        if (cachedImage != null)
                        {
                            vm.ThumbnailPath = cachedImage;
                        }

                        results.Add(vm);
                    }

                    Dispatcher.Invoke(() =>
                    {
                        _molecules.Clear();
                        foreach (MoleculeViewModel vm in results)
                        {
                            _molecules.Add(vm);
                        }

                        moleculeList.Visibility = Visibility.Visible;
                        molStatusText.Text = string.Format("{0} result(s) found.", results.Count);
                        _isSearching = false;
                    });
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        molStatusText.Text = string.Format("Search failed: {0}", ex.Message);
                        _isSearching = false;
                    });
                }
            });
        }

        #endregion

        #region Molecule Selection and 3D Loading

        private void MoleculeList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            MoleculeViewModel selected = moleculeList.SelectedItem as MoleculeViewModel;
            if (selected == null)
            {
                return;
            }

            _selectedMolecule = selected;
            LoadMolecule3D(selected.Cid);

            propertiesPanel.Visibility = Visibility.Visible;
            molNameText.Text = selected.DisplayName;
            molFormulaText.Text = string.Format("Formula: {0}", selected.Formula);
            molWeightText.Text = string.Format("Weight: {0}", selected.Weight);
            molIupacText.Text = selected.IupacName;

            insertSvgButton.IsEnabled = true;
            insertImageButton.IsEnabled = false;
            insert3DButton.IsEnabled = false;
        }

        private void LoadMolecule3D(int cid)
        {
            loadingOverlay.Visibility = Visibility.Visible;
            emptyStateText.Visibility = Visibility.Collapsed;

            Task.Run(() =>
            {
                try
                {
                    PubChemCompoundData compound = null;

                    string cached3D = _cache.GetCached3DJson(cid);
                    if (cached3D != null)
                    {
                        PubChemResponse response = Newtonsoft.Json.JsonConvert.DeserializeObject<PubChemResponse>(cached3D);
                        if (response != null && response.Compounds != null && response.Compounds.Length > 0)
                        {
                            compound = response.Compounds[0];
                        }
                    }
                    else
                    {
                        compound = _pubChemClient.Get3DConformer(cid);
                        if (compound != null)
                        {
                            string json = _pubChemClient.GetRawJson(cid, true);
                            if (json != null)
                            {
                                _cache.Save3DJson(cid, json);
                            }
                        }
                    }

                    PubChemCompoundData compound2D = null;
                    string cached2D = _cache.GetCached2DJson(cid);
                    if (cached2D != null)
                    {
                        PubChemResponse response2D = Newtonsoft.Json.JsonConvert.DeserializeObject<PubChemResponse>(cached2D);
                        if (response2D != null && response2D.Compounds != null && response2D.Compounds.Length > 0)
                        {
                            compound2D = response2D.Compounds[0];
                        }
                    }
                    else
                    {
                        compound2D = _pubChemClient.Get2DStructure(cid);
                        if (compound2D != null)
                        {
                            string json2D = _pubChemClient.GetRawJson(cid, false);
                            if (json2D != null)
                            {
                                _cache.Save2DJson(cid, json2D);
                            }
                        }
                    }

                    if (compound != null)
                    {
                        Model3DGroup model = MoleculeGeometry.BuildMoleculeModel(compound);

                        Dispatcher.Invoke(() =>
                        {
                            _current3DCompound = compound;
                            _current2DCompound = compound2D;

                            moleculeModel.Content = model;
                            AutoFitCamera(compound);

                            loadingOverlay.Visibility = Visibility.Collapsed;
                            insertImageButton.IsEnabled = true;
                            insert3DButton.IsEnabled = true;

                            if (_selectedMolecule != null)
                            {
                                _selectedMolecule.HasConformer3D = true;
                            }

                            molStatusText.Text = "3D structure loaded.";
                        });
                    }
                    else
                    {
                        Dispatcher.Invoke(() =>
                        {
                            _current3DCompound = null;
                            _current2DCompound = compound2D;
                            moleculeModel.Content = null;
                            loadingOverlay.Visibility = Visibility.Collapsed;

                            insertImageButton.IsEnabled = false;
                            insert3DButton.IsEnabled = false;

                            if (_selectedMolecule != null)
                            {
                                _selectedMolecule.HasConformer3D = false;
                            }

                            molStatusText.Text = "No 3D structure available. 2D SVG insert is available.";
                        });
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        loadingOverlay.Visibility = Visibility.Collapsed;
                        molStatusText.Text = string.Format("Failed to load: {0}", ex.Message);
                    });
                }
            });
        }

        private void AutoFitCamera(PubChemCompoundData compound)
        {
            if (compound == null || compound.Coords == null || compound.Coords.Length == 0)
            {
                return;
            }

            PubChemCoords coords = compound.Coords[0];
            if (coords == null || coords.Conformers == null || coords.Conformers.Length == 0)
            {
                return;
            }

            PubChemConformer conf = coords.Conformers[0];
            if (conf.X == null || conf.Y == null || conf.Z == null)
            {
                return;
            }

            double cx = 0;
            double cy = 0;
            double cz = 0;
            double maxDist = 0;
            int count = conf.X.Length;

            for (int i = 0; i < count; i++)
            {
                cx += conf.X[i];
                cy += conf.Y[i];
                cz += conf.Z[i];
            }

            cx /= count;
            cy /= count;
            cz /= count;

            for (int i = 0; i < count; i++)
            {
                double dx = conf.X[i] - cx;
                double dy = conf.Y[i] - cy;
                double dz = conf.Z[i] - cz;
                double dist = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                if (dist > maxDist)
                {
                    maxDist = dist;
                }
            }

            _cameraTarget = new Point3D(cx, cy, cz);
            _cameraDistance = maxDist * 3.5;
            _cameraTheta = 0.3;
            _cameraPhi = Math.PI / 4.0;

            UpdateCameraPosition();
        }

        #endregion

        #region Trackball Camera

        private void Viewport_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _isDragging = true;
            _lastMousePos = e.GetPosition(moleculeViewport);
            moleculeViewport.CaptureMouse();
        }

        private void Viewport_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            _isDragging = false;
            moleculeViewport.ReleaseMouseCapture();
        }

        private void Viewport_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDragging)
            {
                return;
            }

            Point current = e.GetPosition(moleculeViewport);
            double dx = current.X - _lastMousePos.X;
            double dy = current.Y - _lastMousePos.Y;
            _lastMousePos = current;

            _cameraTheta += dx * 0.01;
            _cameraPhi -= dy * 0.01;

            _cameraPhi = Math.Max(0.01, Math.Min(Math.PI - 0.01, _cameraPhi));

            UpdateCameraPosition();
        }

        private void Viewport_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            double factor = e.Delta > 0 ? 0.9 : 1.1;
            _cameraDistance *= factor;
            _cameraDistance = Math.Max(2.0, Math.Min(200.0, _cameraDistance));

            UpdateCameraPosition();
        }

        private void UpdateCameraPosition()
        {
            double x = _cameraTarget.X + _cameraDistance * Math.Sin(_cameraPhi) * Math.Cos(_cameraTheta);
            double y = _cameraTarget.Y + _cameraDistance * Math.Cos(_cameraPhi);
            double z = _cameraTarget.Z + _cameraDistance * Math.Sin(_cameraPhi) * Math.Sin(_cameraTheta);

            camera.Position = new Point3D(x, y, z);
            camera.LookDirection = new Vector3D(
                _cameraTarget.X - x,
                _cameraTarget.Y - y,
                _cameraTarget.Z - z);
        }

        #endregion

        #region Insert Methods

        private void InsertSvg_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedMolecule == null)
            {
                return;
            }

            int cid = _selectedMolecule.Cid;
            PubChemCompoundData compound2D = _current2DCompound;

            if (compound2D == null)
            {
                molStatusText.Text = "No 2D structure data available.";
                return;
            }

            try
            {
                string svg = MoleculeSvgRenderer.RenderSvg(compound2D, 400, 400);
                if (string.IsNullOrEmpty(svg))
                {
                    molStatusText.Text = "Failed to generate SVG.";
                    return;
                }

                string svgPath = _cache.GetSvgPath(cid);
                File.WriteAllText(svgPath, svg);

                InsertPictureOnSlide(svgPath);
                molStatusText.Text = string.Format("Inserted SVG for {0}.", _selectedMolecule.DisplayName);
            }
            catch (Exception ex)
            {
                molStatusText.Text = string.Format("SVG insert failed: {0}", ex.Message);
            }
        }

        private void InsertImage_Click(object sender, RoutedEventArgs e)
        {
            if (_current3DCompound == null)
            {
                return;
            }

            try
            {
                string pngPath = CaptureViewportAsPng();
                if (pngPath == null)
                {
                    molStatusText.Text = "Failed to capture image.";
                    return;
                }

                InsertPictureOnSlide(pngPath);
                molStatusText.Text = string.Format("Inserted PNG for {0}.", _selectedMolecule.DisplayName);
            }
            catch (Exception ex)
            {
                molStatusText.Text = string.Format("PNG insert failed: {0}", ex.Message);
            }
        }

        private void Insert3D_Click(object sender, RoutedEventArgs e)
        {
            if (_current3DCompound == null || _selectedMolecule == null)
            {
                return;
            }

            molStatusText.Text = "Generating 3D model...";

            int cid = _selectedMolecule.Cid;
            PubChemCompoundData compound = _current3DCompound;
            string displayName = _selectedMolecule.DisplayName;

            Task.Run(() =>
            {
                try
                {
                    string glbPath = _cache.GetCachedGlb(cid);
                    if (glbPath == null)
                    {
                        glbPath = _cache.GetGlbPath(cid);
                        MoleculeGlbExporter.ExportGlb(compound, glbPath);
                    }

                    string finalGlbPath = glbPath;

                    Dispatcher.Invoke(() =>
                    {
                        try
                        {
                            PowerPoint.Application app = Globals.ThisAddIn.Application;
                            PowerPoint.Slide slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                            if (slide == null)
                            {
                                molStatusText.Text = "No active slide.";
                                return;
                            }

                            string version = app.Version;
                            bool supportsAdd3DModel = version != null && version.StartsWith("16.");

                            if (supportsAdd3DModel)
                            {
                                try
                                {
                                    dynamic shapes = slide.Shapes;
                                    dynamic shape3D = shapes.Add3DModel(
                                        finalGlbPath,
                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                        50f, 50f, 300f, 300f);

                                    float slideWidth = slide.CustomLayout.Width;
                                    float slideHeight = slide.CustomLayout.Height;
                                    shape3D.Left = (slideWidth - shape3D.Width) / 2f;
                                    shape3D.Top = (slideHeight - shape3D.Height) / 2f;

                                    molStatusText.Text = string.Format("Inserted 3D model for {0}.", displayName);
                                }
                                catch (Exception)
                                {
                                    InsertPictureOnSlide(CaptureViewportAsPng());
                                    molStatusText.Text = string.Format("3D not supported, inserted PNG for {0}.", displayName);
                                }
                            }
                            else
                            {
                                InsertPictureOnSlide(CaptureViewportAsPng());
                                molStatusText.Text = string.Format("Office < 2019, inserted PNG for {0}.", displayName);
                            }
                        }
                        catch (Exception ex)
                        {
                            molStatusText.Text = string.Format("Insert failed: {0}", ex.Message);
                        }
                    });
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        molStatusText.Text = string.Format("GLB generation failed: {0}", ex.Message);
                    });
                }
            });
        }

        #endregion

        #region Private Helpers

        private string CaptureViewportAsPng()
        {
            int width = (int)moleculeViewport.ActualWidth;
            int height = (int)moleculeViewport.ActualHeight;

            if (width < 10 || height < 10)
            {
                width = 400;
                height = 400;
            }

            RenderTargetBitmap rtb = new RenderTargetBitmap(
                width, height, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(moleculeViewport);

            PngBitmapEncoder encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(rtb));

            string tempPath = Path.Combine(Path.GetTempPath(),
                string.Format("biomedpptx_mol_{0}.png", DateTime.Now.Ticks));

            using (FileStream fs = new FileStream(tempPath, FileMode.Create))
            {
                encoder.Save(fs);
            }

            return tempPath;
        }

        private void InsertPictureOnSlide(string imagePath)
        {
            if (string.IsNullOrEmpty(imagePath) || !File.Exists(imagePath))
            {
                return;
            }

            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null)
                {
                    molStatusText.Text = "No active slide.";
                    return;
                }

                PowerPoint.Shape shape = slide.Shapes.AddPicture(
                    imagePath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    0f, 0f);

                float slideWidth = slide.CustomLayout.Width;
                float slideHeight = slide.CustomLayout.Height;
                shape.Left = (slideWidth - shape.Width) / 2f;
                shape.Top = (slideHeight - shape.Height) / 2f;
            }
            catch (Exception ex)
            {
                molStatusText.Text = string.Format("Insert failed: {0}", ex.Message);
            }
        }

        #endregion
    }
}
