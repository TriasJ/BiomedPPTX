using Newtonsoft.Json;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class TilingMetadata
    {
        [JsonProperty("axis")]
        public string Axis { get; set; }

        [JsonProperty("svg_width")]
        public float SvgWidth { get; set; }

        [JsonProperty("svg_height")]
        public float SvgHeight { get; set; }

        [JsonProperty("pptx_width")]
        public float PptxWidth { get; set; }

        [JsonProperty("pptx_height")]
        public float PptxHeight { get; set; }

        [JsonProperty("overlap_px")]
        public float OverlapPx { get; set; }

        [JsonProperty("overlap_pt")]
        public float OverlapPt { get; set; }

        [JsonProperty("overlap_x_px")]
        public float OverlapXPx { get; set; }

        [JsonProperty("overlap_y_px")]
        public float OverlapYPx { get; set; }

        [JsonProperty("overlap_x_pt")]
        public float OverlapXPt { get; set; }

        [JsonProperty("overlap_y_pt")]
        public float OverlapYPt { get; set; }

        [JsonProperty("offset_rows")]
        public bool OffsetRows { get; set; }

        [JsonProperty("offset_amount")]
        public float OffsetAmount { get; set; }
    }
}
