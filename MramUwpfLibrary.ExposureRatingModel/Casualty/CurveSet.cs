using MramUwpfLibrary.Common.Policies;
using TruncatedParetos = MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.TruncatedParetos;
using MixedExponentials = MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty
{
    public abstract class CurveSet
    {
        public string Id { get; set; }
        public double Weight { get; set; }
        public IPolicySet PolicySet { get; set; }
    }

    public class MixedExponentialCurve : CurveSet
    {
        public MixedExponentials.Parameters CurveParameters { get; set; }
        public bool ForceWithinLimits { get; set; }
    }

    public class TruncatedParetoCurve : CurveSet
    {
        public TruncatedParetos.Parameters CurveParameters { get; set; }
    }
}
