using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.ProfessionalLiability
{
    internal class MiscellaneousSubline : BaseSubline
    {
        public MiscellaneousSubline()
        {
            ImageSource = Resources.Misc.ToBitmapSource();
            Code = BexConstants.ProfessionalLiabilityMiscellaneousSublineCode;
            ImageSource.Freeze();
        }
    }
}
