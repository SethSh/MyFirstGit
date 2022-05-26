using SubmissionCollector.Models.DataComponents;

namespace SubmissionCollector.Models.Package.DataComponents
{
    public abstract class BasePackageExcelMatrix : BaseExcelMatrix
    {
        public override string RangeName => $"submission.{ExcelRangeName}";
    }

}
