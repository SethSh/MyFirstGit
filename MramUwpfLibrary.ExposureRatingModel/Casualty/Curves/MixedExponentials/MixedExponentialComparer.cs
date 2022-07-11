using System.Collections.Generic;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials
{
    public class MixedExponentialComparer : IEqualityComparer<MixedExponentialCurve>
    {
        public bool Equals(MixedExponentialCurve x, MixedExponentialCurve y)
        {
            return y != null && x != null && x.Id.Equals(y.Id);
        }
        public int GetHashCode(MixedExponentialCurve x)
        {
            return x.GetHashCode();
        }
    }
}