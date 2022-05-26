using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.RangeSizeModifier
{
    internal abstract class BaseRangeResizer : IRangeResizer
    {
        public ISegmentExcelMatrix ExcelMatrix { get; set; }
        public Range ExcelRange { get; set; }
        
        public abstract void ModifyRange();
        
        public virtual void SetCommonProperties(ISegmentExcelMatrix excelMatrix, Range range)
        { 
            ExcelMatrix = excelMatrix;
            ExcelRange = ExcelMatrix.RangeName.GetRange();
        }

        public abstract bool Validate(ISegmentExcelMatrix excelMatrix, Range range);

        public abstract bool IsOkToReformat();

        public bool IsOkToReformatCommon(string message)
        {
            var sb = new StringBuilder();
            sb.AppendLine(message);
            sb.AppendLine();
            sb.AppendLine("Are you sure you want to proceed?");

            return MessageHelper.ShowWithYesNo(sb.ToString()) == DialogResult.Yes;
        }
    }
}
