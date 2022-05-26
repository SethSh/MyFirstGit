using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.ProfessionalLiability
{
    internal class PublicOfficersErrorAndOmissionsSubline : BaseSubline
    {
        public PublicOfficersErrorAndOmissionsSubline()
        {
            ImageSource = Resources.Official.ToBitmapSource();
            Code = BexConstants.ProfessionalLiabilityPublicOfficersErrorAndOmissionsSublineCode;
            ImageSource.Freeze();
        }
    }
}
