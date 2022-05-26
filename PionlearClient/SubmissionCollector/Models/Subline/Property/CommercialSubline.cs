using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.Property
{
    internal class CommercialSubline : BaseSubline
    {
        public CommercialSubline()
        {
            Code = BexConstants.CommercialPropertySublineCode;
            ImageSource = Resources.Factory.ToBitmapSource();
            ImageSource.Freeze();
            LineOfBusinessType = LineOfBusinessType.Property;
        }

        public override bool HasHazardProfile => false;

    }
}
