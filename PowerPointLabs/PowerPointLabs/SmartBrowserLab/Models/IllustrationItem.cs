using System.Collections.Generic;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class IllustrationItem
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string PptxFile { get; set; }
        public int PptxSlide { get; set; }
        public int PptxShapeIndex { get; set; }
        public string SvgPath { get; set; }
        public string PngPath { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
        public string Topic { get; set; }
        public string SlideTitle { get; set; }
        public List<string> Tags { get; set; }
        public TilingMetadata TilingMeta { get; set; }
    }
}
