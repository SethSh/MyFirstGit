using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.ProfessionalLiability
{
    internal class EmploymentPracticesSubline : BaseSubline
    {
        public EmploymentPracticesSubline()
        {
            Code = BexConstants.ProfessionalLiabilityEmploymentPracticesSublineCode;
            ImageSource = Resources.Briefcase.ToBitmapSource();
            ImageSource.Freeze();
        }
    }
}
