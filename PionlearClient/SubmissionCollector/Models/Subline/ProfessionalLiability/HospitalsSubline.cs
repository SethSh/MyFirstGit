using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.ProfessionalLiability
{
    internal class HospitalsSubline : BaseSubline
    {
        public HospitalsSubline()
        {
            Code = BexConstants.ProfessionalLiabilityHospitalsSublineCode;
            ImageSource = Resources.Hospital.ToBitmapSource();
            ImageSource.Freeze();
        }
    }
}
