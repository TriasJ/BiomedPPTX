using System.Windows.Forms;

namespace PowerPointLabs.SmartBrowserLab
{
    public partial class SmartBrowserPane : UserControl
    {
        public SmartBrowserPane()
        {
            InitializeComponent();
        }

        public void InitBrowser(string dbPath, string assetsBasePath)
        {
            smartBrowserPaneWPF1.Initialize(dbPath, assetsBasePath);
        }
    }
}
