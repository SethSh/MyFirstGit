using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.GeneralLiability
{
    class PremOpsSubline : BaseSubline
    {
        public PremOpsSubline()
        {
            Code = BexConstants.GeneralLiabilityPremOpsSublineCode;
            ImageSource = Resources.WatchYourStep.ToBitmapSource();
            ImageSource.Freeze();
        }
    }
}
