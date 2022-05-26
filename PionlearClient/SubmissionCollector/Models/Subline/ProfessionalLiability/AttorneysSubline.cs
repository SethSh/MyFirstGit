using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.ProfessionalLiability
{
    internal class AttorneysSubline : BaseSubline
    {
        public AttorneysSubline()
        {
            Code = BexConstants.ProfessionalLiabilityAttorneysSublineCode;
            ImageSource = Resources.ScienceScales.ToBitmapSource();
            ImageSource.Freeze();
        }
    }
}
