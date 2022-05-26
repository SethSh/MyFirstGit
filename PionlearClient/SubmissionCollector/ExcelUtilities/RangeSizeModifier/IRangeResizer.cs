using Microsoft.Office.Interop.Excel;
using SubmissionCollector.Models.Segment.DataComponents;

namespace SubmissionCollector.ExcelUtilities.RangeSizeModifier
{
    internal interface IRangeResizer
    {
        void ModifyRange();
        bool Validate(ISegmentExcelMatrix excelMatrix, Range range);
        void SetCommonProperties(ISegmentExcelMatrix excelMatrix, Range range);
        bool IsOkToReformat();
    }
}