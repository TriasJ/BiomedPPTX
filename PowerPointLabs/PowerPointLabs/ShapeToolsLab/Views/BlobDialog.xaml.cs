using System;
using System.Windows;

using PowerPointLabs.ShapeToolsLab.Services;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ShapeToolsLab.Views
{
    public partial class BlobDialog : Window
    {
        private int _seed = 42;

        public BlobDialog()
        {
            InitializeComponent();
        }

        private void ModeChanged(object sender, RoutedEventArgs e)
        {
            if (blobPanel == null || badgePanel == null)
            {
                return;
            }

            if (badgeModeRadio.IsChecked == true)
            {
                blobPanel.Visibility = Visibility.Collapsed;
                badgePanel.Visibility = Visibility.Visible;
                pointCountLabel.Text = "Segments:";
                pointCountSlider.Minimum = 64;
                pointCountSlider.Maximum = 512;
                pointCountSlider.Value = 256;
                pointCountSlider.TickFrequency = 32;
            }
            else
            {
                blobPanel.Visibility = Visibility.Visible;
                badgePanel.Visibility = Visibility.Collapsed;
                pointCountLabel.Text = "Point Count:";
                pointCountSlider.Minimum = 5;
                pointCountSlider.Maximum = 30;
                pointCountSlider.Value = 10;
                pointCountSlider.TickFrequency = 1;
            }
        }

        private void Randomize_Click(object sender, RoutedEventArgs e)
        {
            _seed = new Random().Next(1, 99999);
            seedLabel.Text = "Seed: " + _seed.ToString();
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null)
                {
                    return;
                }

                float cx = app.ActivePresentation.PageSetup.SlideWidth / 2f;
                float cy = app.ActivePresentation.PageSetup.SlideHeight / 2f;
                float radius = (float)radiusSlider.Value;

                if (badgeModeRadio.IsChecked == true)
                {
                    int bumpCount = (int)bumpCountSlider.Value;
                    float bumpHeight = (float)bumpHeightSlider.Value;
                    int segments = (int)pointCountSlider.Value;
                    BlobGenerator.GenerateBadge(slide, cx, cy, radius, bumpHeight, bumpCount, segments);
                }
                else
                {
                    float jitter = (float)jitterSlider.Value;
                    int pointCount = (int)pointCountSlider.Value;
                    BlobGenerator.Generate(slide, cx, cy, radius, jitter, pointCount, _seed);
                }

                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "BiomedPPTX");
            }
        }
    }
}
