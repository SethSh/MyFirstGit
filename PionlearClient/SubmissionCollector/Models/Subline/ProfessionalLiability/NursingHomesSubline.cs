using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.ProfessionalLiability
{
    internal class NursingHomesSubline : BaseSubline
    {
        public NursingHomesSubline()
        {
            Code = BexConstants.ProfessionalLiabilityNursingHomesSublineCode;
            ImageSource = Resources.ElderlyPerson.ToBitmapSource();
            ImageSource.Freeze();
        }
    }
}
