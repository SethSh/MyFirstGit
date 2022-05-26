using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.AutoPhysicalDamage
{
    internal class CommercialSubline : BaseSubline
    {
        public CommercialSubline()
        {
            Code = BexConstants.AutoPhysicalDamageCommercialSublineCode;
            ImageSource = Resources.TruckApd.ToBitmapSource();
            ImageSource.Freeze();
            LineOfBusinessType = LineOfBusinessType.PhysicalDamage;
        }

        public override bool HasHazardProfile => false;
    }
}
