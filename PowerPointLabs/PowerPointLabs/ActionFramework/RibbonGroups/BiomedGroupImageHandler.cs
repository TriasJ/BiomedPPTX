using System.Drawing;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.RibbonGroups
{
    [ExportImageRibbonId(CommonText.BiomedGroupId)]
    class BiomedGroupImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.BiomedGroup);
        }
    }
}
