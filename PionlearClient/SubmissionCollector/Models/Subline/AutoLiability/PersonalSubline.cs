using System;
using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.AutoLiability
{
    [Serializable]
    internal class PersonalSubline : BaseSubline
    {
        public PersonalSubline()
        {
            ImageSource = Resources.Car.ToBitmapSource();
            Code = BexConstants.AutoLiabilityPersonalSublineCode;
            ImageSource.Freeze();
        }
    }
}
