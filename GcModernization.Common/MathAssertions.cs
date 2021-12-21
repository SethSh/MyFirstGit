using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GcModernization.Common;

public class MathAssertions
{
    private const double DoubleTolerance = 1e-6;

    public static void IsSameLength<T>(IEnumerable<T> list1, IEnumerable<T> list2)
    {
        Assert.IsTrue(list1.Count() == list2.Count(), "List lengths don't match");
    }

    public static void IsEpsilonEqual(IList<double> list1, IList<double> list2)
    {
        IsSameLength(list1, list2);

        var row = 0;
        foreach (var item1 in list1)
        {
            var item2 = list2[row];
            var isEpsilonEqual = IsEpsilonEqual(item1, item2);
            Assert.IsTrue(isEpsilonEqual, $"Row {row}: {item1} doesn't equal {item2}");
            row++;
        }
    }

    private static bool IsEpsilonEqual(double d1, double d2)
    {
        return Math.Abs(d1 - d2) < DoubleTolerance;
    }
}