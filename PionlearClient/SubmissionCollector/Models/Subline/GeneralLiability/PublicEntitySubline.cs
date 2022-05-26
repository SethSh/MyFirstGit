using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.GeneralLiability
{
    internal class PublicEntitySubline : BaseSubline
    {
        public PublicEntitySubline()
        {
            ImageSource = Resources.Capitol.ToBitmapSource();
            Code = BexConstants.GeneralLiabilityPublicEntitySublineCode;
            ImageSource.Freeze();
        }
    }
}
