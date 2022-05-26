using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.ProfessionalLiability
{
    internal class DirectorsAndOfficersSubline : BaseSubline
    {
        public DirectorsAndOfficersSubline()
        {
            Code = BexConstants.ProfessionalLiabilityDirectorsAndOfficersSublineCode;
            ImageSource = Resources.Suit.ToBitmapSource();
            ImageSource.Freeze();
        }
    }
}
