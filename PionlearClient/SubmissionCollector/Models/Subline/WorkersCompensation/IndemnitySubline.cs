using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.WorkersCompensation
{
    internal class IndemnitySubline : BaseSubline
    {
        public IndemnitySubline()
        {
            Code = BexConstants.WorkersCompensationIndemnitySublineCode;
            ImageSource = Resources.Paycheque.ToBitmapSource();
            ImageSource.Freeze();
            LineOfBusinessType = LineOfBusinessType.WorkersCompensation;
        }

        public override bool HasHazardProfile => false;
    }
}
