using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.GeneralLiability
{
    internal class HomeownersLiabilitySubline : BaseSubline
    {
        public HomeownersLiabilitySubline()
        {
            ImageSource = Resources.House.ToBitmapSource();
            Code = BexConstants.GeneralLiabilityHomeownersLiabilitySublineCode;
            ImageSource.Freeze();
        }
    }
}
