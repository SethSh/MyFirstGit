using System;

namespace PionlearClient
{
    public static class ComparisonExtensions
    {
        private const double DefaultEpsilon = 1e-6;
        
        public static bool IsEqual(this double d1, double d2)
        {
            return d1.IsEpsilonEqual(d2, 0);
        }

        public static bool IsEpsilonEqualIncludingNullAndNaN(this double? d1, double? d2)
        {
            if (d1 == null ^ d2 == null) return false;
            if (d1 == null) return true;

            if (double.IsNaN(d1.Value) ^ double.IsNaN(d2.Value)) return false;
            if (double.IsNaN(d1.Value) && double.IsNaN(d2.Value)) return true;

            return d1.Value.IsEpsilonEqual(d2.Value);
        }

        public static bool IsDateEqualIncludingNull(this DateTimeOffset? d1, DateTimeOffset? d2)
        {
            if (d1 == null ^ d2 == null) return false;
            return d1 == null || d1.Value.DateTime.ToShortDateString() == d2.Value.DateTime.ToShortDateString();
        }

        public static bool IsEpsilonEqual(this double d1, double d2, double eps)
        {
            //both NaN should fail
            if (double.IsNaN(d1) || double.IsNaN(d2))
                return false;

            return Math.Abs(d1 - d2) <= eps;
        }

        public static bool IsEpsilonEqualToOne(this double d1)
        {
            return d1.IsEpsilonEqual(1d, DefaultEpsilon);
        }

        public static bool IsEpsilonEqual(this double d1, double d2)
        {
            return d1.IsEpsilonEqual(d2, DefaultEpsilon);
        }

        public static bool IsEqualTo(this long? first, long? second)
        {
            if (first.HasValue ^ second.HasValue) return false;

            //both null are a match
            if (!first.HasValue) return true;

            return first.Value == second.Value;
        }

    }
}
