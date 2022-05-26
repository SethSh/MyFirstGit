using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.AutoPhysicalDamage
{
    internal class PersonalSubline : BaseSubline
    {
        public PersonalSubline()
        {
            Code = BexConstants.AutoPhysicalDamagePersonalSublineCode;
            ImageSource = Resources.CarApd.ToBitmapSource();
            ImageSource.Freeze();
            LineOfBusinessType = LineOfBusinessType.PhysicalDamage;
        }

        public override bool HasHazardProfile => false;
    }
}
