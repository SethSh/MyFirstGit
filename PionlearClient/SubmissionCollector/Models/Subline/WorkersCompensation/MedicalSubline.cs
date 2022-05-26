using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.WorkersCompensation
{
    internal class MedicalSubline : BaseSubline
    {
        public MedicalSubline()
        {
            Code = BexConstants.WorkersCompensationMedicalSublineCode;
            ImageSource = Resources.WheelBed.ToBitmapSource();
            ImageSource.Freeze();
            LineOfBusinessType = LineOfBusinessType.WorkersCompensation;
        }

        public override bool HasHazardProfile => false;
    }
}
