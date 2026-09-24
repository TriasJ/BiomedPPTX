using System.Windows.Forms;

namespace PowerPointLabs.PatternBrushLab
{
    public partial class PatternBrushPane : UserControl
    {
        public PatternBrushPane()
        {
            InitializeComponent();
        }

        public void InitBrush(string dbPath, string assetsBasePath)
        {
            patternBrushPaneWPF1.Initialize(dbPath, assetsBasePath);
        }
    }
}
