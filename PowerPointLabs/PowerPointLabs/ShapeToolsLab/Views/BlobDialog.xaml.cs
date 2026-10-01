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
                float jitter = (float)jitterSlider.Value;
                int pointCount = (int)pointCountSlider.Value;

                BlobGenerator.Generate(slide, cx, cy, radius, jitter, pointCount, _seed);
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "BiomedPPTX");
            }
        }
    }
}
