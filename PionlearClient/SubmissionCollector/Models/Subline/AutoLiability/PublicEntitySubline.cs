using System;
using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.AutoLiability
{
    [Serializable]
    internal class PublicEntitySubline : BaseSubline
    {
        public PublicEntitySubline()
        {
            ImageSource = Resources.Bus.ToBitmapSource();
            Code = BexConstants.AutoLiabilityPublicEntitySublineCode;
            ImageSource.Freeze();
        }
    }
}
