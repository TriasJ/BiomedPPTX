using System;
using System.Windows;

using PowerPointLabs.ShapeToolsLab.Services;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ShapeToolsLab.Views
{
    public partial class WavyLineDialog : Window
    {
        public WavyLineDialog()
        {
            InitializeComponent();
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

                float slideW = app.ActivePresentation.PageSetup.SlideWidth;
                float slideH = app.ActivePresentation.PageSetup.SlideHeight;
                float amplitude = (float)amplitudeSlider.Value;
                float frequency = (float)frequencySlider.Value;
                float width = (float)widthSlider.Value;
                int points = (int)pointsSlider.Value;
                float startX = (slideW - width) / 2f;
                float startY = slideH / 2f;

                WavyLineGenerator.Generate(slide, startX, startY, width, amplitude, frequency, points);
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "BiomedPPTX");
            }
        }
    }
}
