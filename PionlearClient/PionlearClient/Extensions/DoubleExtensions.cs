using System;

namespace PionlearClient.Extensions
{
    internal static class DoubleExtensions
    {
        internal static bool IsGreaterThan(this double? firstDouble, double? secondDouble)
        {
            if (firstDouble.HasValue ^ secondDouble.HasValue) return false;
            if (!firstDouble.HasValue) return false;

            return firstDouble.Value.IsGreaterThan(secondDouble.Value);
        }

        public static bool IsGreaterThan(this double firstDouble, double? secondDouble)
        {
            return secondDouble.HasValue && firstDouble.IsGreaterThan(secondDouble.Value);
        }

        public static long RoundToLong(this double d)
        {
            return Convert.ToInt64(Math.Round(d, 0, MidpointRounding.AwayFromZero));
        }

        private static bool IsGreaterThan(this double firstDouble, double secondDouble)
        {
            return firstDouble > secondDouble;
        }
        
    }
}
