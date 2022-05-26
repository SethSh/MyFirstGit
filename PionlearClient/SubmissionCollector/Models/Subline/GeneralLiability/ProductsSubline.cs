using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.GeneralLiability
{
    class ProductsSubline : BaseSubline
    {
        public ProductsSubline()
        {
            Code = BexConstants.GeneralLiabilityProductsSublineCode;
            ImageSource = Resources.Ladder.ToBitmapSource();
            ImageSource.Freeze();
        }
    }
}
