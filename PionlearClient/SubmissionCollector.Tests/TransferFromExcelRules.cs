using Microsoft.VisualStudio.TestTools.UnitTesting;
using SubmissionCollector.ExcelUtilities;

namespace SubmissionCollector.Tests
{
    [TestClass]
    public class TransferFromExcelRules
    {
        [TestMethod]
        public void TransferFromExcel_AllDoubles()
        {
            var likeFromExcel = new object[,]
			{
				{1d, 2d, 3d},
				{40.4, 50d, 66d}
			};

            var actual = likeFromExcel.ForceContentToDoubles();
            var expected = new[,]
			{
				{1, 2, 3},
				{40.4, 50, 66}
			};

            MathAssertions.AreEpsilonEquals(expected, actual);
        }

        
        [TestMethod]
        public void TransferFromExcel_DoublesAndStrings()
        {
            var likeFromExcel = new object[,]
			{
				{1d, 2d, "Joe"},
				{40.4, 50d, "Seth"}
			};

            var actual = likeFromExcel.ForceContentToDoubles();
            var expected = new[,]
			{
				{1, 2, double.NaN},
				{40.4, 50, double.NaN}
			};

            MathAssertions.AreEpsilonEquals(expected, actual);
        }
        

        [TestMethod]
        public void TransferFromExcel_DoublesAndIntegers()
        {
            var likeFromExcel = new object[,]
			{
				{1d, 2d, 100},
				{40.4, 50d, 66}
			};

            var actual = likeFromExcel.ForceContentToDoubles();
            var expected = new[,]
			{
				{1, 2, double.NaN},
				{40.4, 50, double.NaN}
			};

            MathAssertions.AreEpsilonEquals(expected, actual);
        }

        
    }
}
