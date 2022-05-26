using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SubmissionCollector.Tests
{
    public static class MathAssertions
    {
        private const double DefaultEpsilon = 1E-6;

        public static bool AreEpsilonEquals(double[,] array1, double[,] array2)
        {
            return AreEpsilonEquals(array1, array2, DefaultEpsilon);
        }

        public static bool AreEpsilonEquals(double[,] array1, double[,] array2, double epsilon)
        {
            AssertEqualSizes(array1, array2);

            var rowCount = array1.GetLength(0);
            var columnCount = array1.GetLength(0);

            for (var row = 0; row < rowCount; row++)
            {
                for (var column = 0; column < columnCount; column++)
                {
                    var expected = array1[row, column];
                    var actual = array2[row, column];

                    if (!double.IsNaN(expected) && !double.IsNaN(actual))
                    {
                        if (Math.Abs(expected - actual) > epsilon)
                        {
                            Assert.Fail($"Expected <{expected}>, Actual <{actual}> in <{row} x {column}>");
                        }
                    }
                    else if (double.IsNaN(expected) ^ double.IsNaN(actual))
                    {
                        Assert.Fail($"Expected <{expected}> or Actual <{actual}> is NaN");
                    }
                }
            }

            return true;
        }

        public static void AssertEqualSizes(double[,] array1, double[,] array2)
        {
            var expectedRowCount = array1.GetLength(0);
            var actualRowCount = array2.GetLength(0);

            if (expectedRowCount != actualRowCount)
            {
                Assert.Fail("Expected <{0}> rows, Actual <{1}> rows", expectedRowCount, actualRowCount);
            }

            var expectedColumnCount = array1.GetLength(1);
            var actualColumnCount = array2.GetLength(1);

            if (expectedColumnCount != actualColumnCount)
            {
                Assert.Fail("Expected <{0}> cols, Actual <{1}> cols", expectedColumnCount, actualColumnCount);
            }
        }
    }
}
