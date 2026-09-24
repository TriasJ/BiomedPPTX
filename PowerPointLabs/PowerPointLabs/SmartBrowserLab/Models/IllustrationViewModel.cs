using System.ComponentModel;

namespace PowerPointLabs.SmartBrowserLab.Views
{
    public class IllustrationViewModel : INotifyPropertyChanged
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string Description { get; set; }
        public string ThumbnailPath { get; set; }
        public string SvgPath { get; set; }
        public string PptxFile { get; set; }
        public int PptxSlide { get; set; }
        public int PptxShapeIndex { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
        public string Topic { get; set; }
        public bool IsTileable { get; set; }
        public string Source { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(string name)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }
    }
}
