using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.ProfessionalLiability
{
    internal class PhysiciansSubline : BaseSubline
    {
        public PhysiciansSubline()
        {
            ImageSource = Resources.Doctor.ToBitmapSource();
            Code = BexConstants.ProfessionalLiabilityPhysiciansSublineCode;
            ImageSource.Freeze();
        }
    }
}
