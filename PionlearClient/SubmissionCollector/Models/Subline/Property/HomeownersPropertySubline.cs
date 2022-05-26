using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.Property
{
    internal class HomeownersPropertySubline : BaseSubline
    {
        public HomeownersPropertySubline()
        {
            Code = BexConstants.GeneralLiabilityHomeownersPropertySublineCode;
            ImageSource = Resources.HouseBlue.ToBitmapSource();
            ImageSource.Freeze();
            LineOfBusinessType = LineOfBusinessType.Property;
        }

        public override bool HasHazardProfile => false;
    }
}
