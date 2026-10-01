using System;
using System.Windows;

using PowerPointLabs.ShapeToolsLab.Services;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ShapeToolsLab.Views
{
    public partial class GearStarDialog : Window
    {
        public GearStarDialog()
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

                float cx = app.ActivePresentation.PageSetup.SlideWidth / 2f;
                float cy = app.ActivePresentation.PageSetup.SlideHeight / 2f;
                float outerR = (float)outerRadiusSlider.Value;
                float innerR = (float)innerRadiusSlider.Value;
                int points = (int)pointsSlider.Value;

                if (starMode.IsChecked == true)
                {
                    GearStarGenerator.GenerateStar(slide, cx, cy, outerR, innerR, points);
                }
                else
                {
                    GearStarGenerator.GenerateGear(slide, cx, cy, outerR, innerR, points);
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
