using System.Collections.Generic;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class BioArtItem
    {
        public int Id { get; set; }
        public string BioArtId { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string Url { get; set; }
        public string License { get; set; }
        public List<string> Keywords { get; set; }
        public List<BioArtFileEndpoint> FileEndpoints { get; set; }
        public bool HasFiles { get; set; }
    }
}
