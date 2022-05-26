using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.ProfessionalLiability
{
    internal class SurgeonsSubline : BaseSubline
    {
        public SurgeonsSubline()
        {
            ImageSource = Resources.Surgeon.ToBitmapSource();
            Code = BexConstants.ProfessionalLiabilitySurgeonsSublineCode;
            ImageSource.Freeze();
        }
    }
}
