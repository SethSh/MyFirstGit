using System;
using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.AutoLiability
{
    [Serializable]
    internal class CommercialSubline : BaseSubline
    {
        public CommercialSubline()
        {
            Code = BexConstants.AutoLiabilityCommercialSublineCode;
            ImageSource = Resources.Truck.ToBitmapSource();
            ImageSource.Freeze();
        }
    }
}
