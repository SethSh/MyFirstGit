using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.Common.ReinsurancePerspectives;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves
{
    public struct CurveInputs
    {
        public double TopLimit { get; set; }
        public double BottomLimit { get; set; }
        public IReinsurancePerspectiveHandler ReinsurancePerspective { get; set; }
        public ReinsuranceAlaeTreatmentType ReinsuranceAlaeTreatment { get; set; }
        public PolicyAlaeTreatmentType PolicyAlaeTreatment { get; set; }
        public double AlaeAdjustmentFactor { get; set; }
    }
}
