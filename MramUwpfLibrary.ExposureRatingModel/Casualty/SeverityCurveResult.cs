
using System.Collections.Generic;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty
{
    public class SeverityCurveResult
    {
        public SeverityCurveResult()
        {
            States = new List<long>();
        }
        public long CurveId { get; set; }
        public double Weight { get; set; }
        public Curves.MixedExponentials.Parameters MixedExponentialCurveParameters { get; set; }
        public Curves.TruncatedParetos.Parameters TruncatedParetoCurveParameters { get; set; }
        public List<long>  States { get; set; }
    }
}
