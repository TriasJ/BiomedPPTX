using System.ComponentModel;

namespace PowerPointLabs.PatternBrushLab.Views
{
    public class PatternViewModel : INotifyPropertyChanged
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public string DisplayName { get; set; }

        public string ThumbnailPath { get; set; }

        public string PptxFile { get; set; }

        public int PptxSlide { get; set; }

        public int PptxShapeIndex { get; set; }

        public string Axis { get; set; }

        public float DefaultOverlap { get; set; }

        public float TileWidth { get; set; }

        public float TileHeight { get; set; }

        public bool OffsetRows { get; set; }

        public float OffsetAmount { get; set; }

        public float SvgWidth { get; set; }

        public string BorderColor { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string name)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }
    }
}
