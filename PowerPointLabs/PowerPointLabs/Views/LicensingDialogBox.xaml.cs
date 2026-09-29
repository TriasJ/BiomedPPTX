using System;
using System.Windows;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Views
{
    public partial class LicensingDialogBox : Window
    {
        private const string SmartAttribution =
            "Illustrations adapted from SMART - Servier Medical ART " +
            "(https://smart.servier.com/), licensed under " +
            "Creative Commons Attribution 3.0 Unported (CC BY 3.0).";

        private const string BioArtAttribution =
            "Illustrations from NCBI BioArt, NIAID Visual & Medical Arts " +
            "(https://bioart.niaid.nih.gov/), Public Domain / CC-BY.";

        private const string FullAttribution =
            "Licensing & Attribution\n\n" +
            "BiomedPPTX\n" +
            "License: GNU General Public License v2.0 (GPLv2)\n" +
            "https://github.com/TriasJ/BiomedPPTX\n\n" +
            "Based on PowerPointLabs\n" +
            "License: GPLv2 - School of Computing, NUS\n" +
            "https://github.com/PowerPointLabs/PowerPointLabs\n\n" +
            "SMART - Servier Medical ART\n" +
            "License: Creative Commons Attribution 3.0 Unported (CC BY 3.0)\n" +
            "Copyright: Les Laboratoires Servier\n" +
            "https://smart.servier.com/\n\n" +
            "NCBI BioArt\n" +
            "License: Public Domain / CC-BY\n" +
            "Copyright: NIAID Visual & Medical Arts\n" +
            "https://bioart.niaid.nih.gov/";

        public LicensingDialogBox()
        {
            InitializeComponent();
        }

        private void CopyAttribution_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Windows.Clipboard.SetText(SmartAttribution + "\n\n" + BioArtAttribution);
                MessageBox.Show("Attribution text copied to clipboard.", "BiomedPPTX",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception)
            {
            }
        }

        private void AddAttributionSlide_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation pres = app.ActivePresentation;

                int slideCount = pres.Slides.Count;
                PowerPoint.Slide slide = pres.Slides.Add(slideCount + 1,
                    PowerPoint.PpSlideLayout.ppLayoutBlank);

                float slideWidth = pres.PageSetup.SlideWidth;
                float slideHeight = pres.PageSetup.SlideHeight;

                PowerPoint.Shape titleBox = slide.Shapes.AddTextbox(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    40, 20, slideWidth - 80, 50);
                titleBox.TextFrame.TextRange.Text = "Licensing & Attribution";
                titleBox.TextFrame.TextRange.Font.Size = 28;
                titleBox.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                titleBox.TextFrame.TextRange.Font.Color.RGB = 0xC07000;

                string bodyText =
                    "SMART - Servier Medical ART\n" +
                    SmartAttribution + "\n\n" +
                    "NCBI BioArt\n" +
                    BioArtAttribution + "\n\n" +
                    "BiomedPPTX is based on PowerPointLabs (GPLv2)\n" +
                    "School of Computing, National University of Singapore\n" +
                    "https://github.com/PowerPointLabs/PowerPointLabs\n\n" +
                    "BiomedPPTX (GPLv2)\n" +
                    "https://github.com/TriasJ/BiomedPPTX";

                PowerPoint.Shape bodyBox = slide.Shapes.AddTextbox(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    40, 80, slideWidth - 80, slideHeight - 120);
                bodyBox.TextFrame.TextRange.Text = bodyText;
                bodyBox.TextFrame.TextRange.Font.Size = 12;
                bodyBox.TextFrame.TextRange.Font.Color.RGB = 0x444444;
                bodyBox.TextFrame.WordWrap = Microsoft.Office.Core.MsoTriState.msoTrue;

                pres.Application.ActiveWindow.View.GotoSlide(slideCount + 1);

                MessageBox.Show("Attribution slide added at the end of the presentation.",
                    "BiomedPPTX", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to add attribution slide: " + ex.Message,
                    "BiomedPPTX", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
